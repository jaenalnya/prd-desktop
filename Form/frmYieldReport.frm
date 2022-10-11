VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LvButton.ocx"
Object = "{8E048CF2-F435-45C9-8A6F-4646F9E1B5F4}#1.0#0"; "XPTab.ocx"
Begin VB.Form frmYieldReport 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Yield Report"
   ClientHeight    =   11565
   ClientLeft      =   5355
   ClientTop       =   1785
   ClientWidth     =   15690
   Icon            =   "frmYieldReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11565
   ScaleWidth      =   15690
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   6795
      Top             =   10710
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   7245
      Top             =   10710
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
      Left            =   135
      TabIndex        =   599
      Top             =   10530
      Width           =   6630
      Begin VB.Shape Shape7 
         BackColor       =   &H0080FF80&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   270
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
         TabIndex        =   603
         Top             =   315
         Width           =   510
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   1575
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
         TabIndex        =   602
         Top             =   315
         Width           =   1230
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00C000C0&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   3330
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
         TabIndex        =   601
         Top             =   315
         Width           =   1455
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   5265
         Top             =   270
         Width           =   240
      End
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
         TabIndex        =   600
         Top             =   315
         Width           =   825
      End
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
      Left            =   8730
      TabIndex        =   594
      Top             =   10530
      Width           =   6630
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
         TabIndex        =   598
         Top             =   315
         Width           =   825
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   5265
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
         TabIndex        =   597
         Top             =   315
         Width           =   1455
      End
      Begin VB.Shape Shape12 
         BackColor       =   &H00C000C0&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   3375
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
         TabIndex        =   596
         Top             =   315
         Width           =   1230
      End
      Begin VB.Shape Shape13 
         BackColor       =   &H0000FFFF&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   1575
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
      Begin VB.Shape Shape14 
         BackColor       =   &H0080FF80&
         BackStyle       =   1  'Opaque
         Height          =   240
         Left            =   270
         Top             =   270
         Width           =   240
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   7740
      Top             =   10710
   End
   Begin prjXTab.XTab XTab1 
      Height          =   9870
      Left            =   135
      TabIndex        =   0
      Top             =   630
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   17410
      TabCaption(0)   =   "MESIN 1 - 20"
      TabContCtrlCnt(0)=   20
      Tab(0)ContCtrlCap(1)=   "frameDesign0"
      Tab(0)ContCtrlCap(2)=   "frameDesign2"
      Tab(0)ContCtrlCap(3)=   "frameDesign3"
      Tab(0)ContCtrlCap(4)=   "frameDesign4"
      Tab(0)ContCtrlCap(5)=   "frameDesign5"
      Tab(0)ContCtrlCap(6)=   "frameDesign6"
      Tab(0)ContCtrlCap(7)=   "frameDesign7"
      Tab(0)ContCtrlCap(8)=   "frameDesign8"
      Tab(0)ContCtrlCap(9)=   "frameDesign9"
      Tab(0)ContCtrlCap(10)=   "frameDesign10"
      Tab(0)ContCtrlCap(11)=   "frameDesign11"
      Tab(0)ContCtrlCap(12)=   "frameDesign12"
      Tab(0)ContCtrlCap(13)=   "frameDesign13"
      Tab(0)ContCtrlCap(14)=   "frameDesign14"
      Tab(0)ContCtrlCap(15)=   "frameDesign15"
      Tab(0)ContCtrlCap(16)=   "frameDesign16"
      Tab(0)ContCtrlCap(17)=   "frameDesign17"
      Tab(0)ContCtrlCap(18)=   "frameDesign18"
      Tab(0)ContCtrlCap(19)=   "frameDesign19"
      Tab(0)ContCtrlCap(20)=   "frameDesign20"
      TabCaption(1)   =   "MESIN 21 - 40"
      TabContCtrlCnt(1)=   20
      Tab(1)ContCtrlCap(1)=   "frameDesign26"
      Tab(1)ContCtrlCap(2)=   "frameDesign27"
      Tab(1)ContCtrlCap(3)=   "frameDesign28"
      Tab(1)ContCtrlCap(4)=   "frameDesign29"
      Tab(1)ContCtrlCap(5)=   "frameDesign30"
      Tab(1)ContCtrlCap(6)=   "frameDesign31"
      Tab(1)ContCtrlCap(7)=   "frameDesign32"
      Tab(1)ContCtrlCap(8)=   "frameDesign33"
      Tab(1)ContCtrlCap(9)=   "frameDesign34"
      Tab(1)ContCtrlCap(10)=   "frameDesign35"
      Tab(1)ContCtrlCap(11)=   "frameDesign36"
      Tab(1)ContCtrlCap(12)=   "frameDesign37"
      Tab(1)ContCtrlCap(13)=   "frameDesign38"
      Tab(1)ContCtrlCap(14)=   "frameDesign39"
      Tab(1)ContCtrlCap(15)=   "frameDesign40"
      Tab(1)ContCtrlCap(16)=   "frameDesign21"
      Tab(1)ContCtrlCap(17)=   "frameDesign22"
      Tab(1)ContCtrlCap(18)=   "frameDesign23"
      Tab(1)ContCtrlCap(19)=   "frameDesign24"
      Tab(1)ContCtrlCap(20)=   "frameDesign25"
      TabCaption(2)   =   "MESIN 41 - 60"
      TabContCtrlCnt(2)=   5
      Tab(2)ContCtrlCap(1)=   "frameDesign41"
      Tab(2)ContCtrlCap(2)=   "frameDesign1"
      Tab(2)ContCtrlCap(3)=   "frameDesign42"
      Tab(2)ContCtrlCap(4)=   "frameDesign43"
      Tab(2)ContCtrlCap(5)=   "frameDesign44"
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
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   41
         Left            =   -74685
         TabIndex        =   578
         Top             =   540
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
            Index           =   41
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   585
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
            Index           =   41
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   584
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
            Index           =   41
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   583
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
            Index           =   41
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   582
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
            Index           =   41
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   581
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
            Index           =   41
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   580
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
            Index           =   41
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   579
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
            Index           =   5
            Left            =   1935
            TabIndex        =   591
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
            Index           =   6
            Left            =   1035
            TabIndex        =   590
            Top             =   675
            Width           =   645
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
            TabIndex        =   589
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
            Index           =   7
            Left            =   90
            TabIndex        =   588
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
            Index           =   8
            Left            =   90
            TabIndex        =   587
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
            Index           =   9
            Left            =   90
            TabIndex        =   586
            Top             =   1800
            Width           =   870
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   1
         Left            =   -68835
         TabIndex        =   577
         Top             =   540
         Width           =   2760
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   42
         Left            =   -71760
         TabIndex        =   563
         Top             =   540
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
            Index           =   42
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   570
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
            Index           =   42
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   569
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
            Index           =   42
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   568
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
            Index           =   42
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   567
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
            Index           =   42
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   566
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
            Index           =   42
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   565
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
            Index           =   42
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   564
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
            Index           =   205
            Left            =   90
            TabIndex        =   576
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
            Index           =   206
            Left            =   90
            TabIndex        =   575
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
            Index           =   207
            Left            =   90
            TabIndex        =   574
            Top             =   990
            Width           =   960
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
            TabIndex        =   573
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
            Index           =   208
            Left            =   1035
            TabIndex        =   572
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
            Index           =   209
            Left            =   1935
            TabIndex        =   571
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   43
         Left            =   -65910
         TabIndex        =   562
         Top             =   540
         Width           =   2760
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   44
         Left            =   -62985
         TabIndex        =   561
         Top             =   540
         Width           =   2760
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   26
         Left            =   -74685
         TabIndex        =   547
         Top             =   2880
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
            Index           =   26
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   554
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
            Index           =   26
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   553
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
            Index           =   26
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   552
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
            Index           =   26
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   551
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
            Index           =   26
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   550
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
            Index           =   26
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   549
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
            Index           =   26
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   548
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
            Index           =   134
            Left            =   1935
            TabIndex        =   560
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
            Index           =   133
            Left            =   1035
            TabIndex        =   559
            Top             =   675
            Width           =   645
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
            TabIndex        =   558
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
            Index           =   132
            Left            =   90
            TabIndex        =   557
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
            Index           =   131
            Left            =   90
            TabIndex        =   556
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
            Index           =   130
            Left            =   90
            TabIndex        =   555
            Top             =   1800
            Width           =   870
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   27
         Left            =   -71760
         TabIndex        =   533
         Top             =   2880
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
            Index           =   27
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   540
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
            Index           =   27
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   539
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
            Index           =   27
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   538
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
            Index           =   27
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   537
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
            Index           =   27
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   536
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
            Index           =   27
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   535
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
            Index           =   27
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   534
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
            Index           =   139
            Left            =   90
            TabIndex        =   546
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
            Index           =   138
            Left            =   90
            TabIndex        =   545
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
            Index           =   137
            Left            =   90
            TabIndex        =   544
            Top             =   990
            Width           =   960
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
            TabIndex        =   543
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
            Index           =   136
            Left            =   1035
            TabIndex        =   542
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
            Index           =   135
            Left            =   1935
            TabIndex        =   541
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   28
         Left            =   -68835
         TabIndex        =   519
         Top             =   2880
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
            Index           =   28
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   526
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
            Index           =   28
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   525
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
            Index           =   28
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   524
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
            Index           =   28
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   523
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
            Index           =   28
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   522
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
            Index           =   28
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   521
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
            Index           =   28
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   520
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
            Index           =   144
            Left            =   90
            TabIndex        =   532
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
            Index           =   143
            Left            =   90
            TabIndex        =   531
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
            Index           =   142
            Left            =   90
            TabIndex        =   530
            Top             =   990
            Width           =   960
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
            TabIndex        =   529
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
            Index           =   141
            Left            =   1035
            TabIndex        =   528
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
            Index           =   140
            Left            =   1935
            TabIndex        =   527
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   29
         Left            =   -65910
         TabIndex        =   505
         Top             =   2880
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
            Index           =   29
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   512
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
            Index           =   29
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   511
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
            Index           =   29
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   510
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
            Index           =   29
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   509
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
            Index           =   29
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   508
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
            Index           =   29
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   507
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
            Index           =   29
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   506
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
            Index           =   149
            Left            =   90
            TabIndex        =   518
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
            Index           =   148
            Left            =   90
            TabIndex        =   517
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
            Index           =   147
            Left            =   90
            TabIndex        =   516
            Top             =   990
            Width           =   960
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
            TabIndex        =   515
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
            Index           =   146
            Left            =   1035
            TabIndex        =   514
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
            Index           =   145
            Left            =   1935
            TabIndex        =   513
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   30
         Left            =   -62985
         TabIndex        =   491
         Top             =   2880
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
            Index           =   30
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   498
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
            Index           =   30
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   497
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
            Index           =   30
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   496
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
            Index           =   30
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   495
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
            Index           =   30
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   494
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
            Index           =   30
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   493
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
            Index           =   30
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   492
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
            Index           =   154
            Left            =   90
            TabIndex        =   504
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
            Index           =   153
            Left            =   90
            TabIndex        =   503
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
            Index           =   152
            Left            =   90
            TabIndex        =   502
            Top             =   990
            Width           =   960
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
            TabIndex        =   501
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
            Index           =   151
            Left            =   1035
            TabIndex        =   500
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
            Index           =   150
            Left            =   1935
            TabIndex        =   499
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   31
         Left            =   -74685
         TabIndex        =   477
         Top             =   5220
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
            Index           =   31
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   484
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
            Index           =   31
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   483
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
            Index           =   31
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   482
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
            Index           =   31
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   481
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
            Index           =   31
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   480
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
            Index           =   31
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   479
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
            Index           =   31
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   478
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
            Index           =   159
            Left            =   1935
            TabIndex        =   490
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
            Index           =   158
            Left            =   1035
            TabIndex        =   489
            Top             =   675
            Width           =   645
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
            TabIndex        =   488
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
            Index           =   157
            Left            =   90
            TabIndex        =   487
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
            Index           =   156
            Left            =   90
            TabIndex        =   486
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
            Index           =   155
            Left            =   90
            TabIndex        =   485
            Top             =   1800
            Width           =   870
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   32
         Left            =   -71760
         TabIndex        =   463
         Top             =   5220
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
            Index           =   32
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   470
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
            Index           =   32
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   469
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
            Index           =   32
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   468
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
            Index           =   32
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   467
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
            Index           =   32
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   466
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
            Index           =   32
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   465
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
            Index           =   32
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   464
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
            Index           =   164
            Left            =   90
            TabIndex        =   476
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
            Index           =   163
            Left            =   90
            TabIndex        =   475
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
            Index           =   162
            Left            =   90
            TabIndex        =   474
            Top             =   990
            Width           =   960
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
            TabIndex        =   473
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
            Index           =   161
            Left            =   1035
            TabIndex        =   472
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
            Index           =   160
            Left            =   1935
            TabIndex        =   471
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   33
         Left            =   -68835
         TabIndex        =   449
         Top             =   5220
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
            Index           =   33
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   456
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
            Index           =   33
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   455
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
            Index           =   33
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   454
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
            Index           =   33
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   453
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
            Index           =   33
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   452
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
            Index           =   33
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   451
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
            Index           =   33
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   450
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
            Index           =   169
            Left            =   90
            TabIndex        =   462
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
            Index           =   168
            Left            =   90
            TabIndex        =   461
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
            Index           =   167
            Left            =   90
            TabIndex        =   460
            Top             =   990
            Width           =   960
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
            TabIndex        =   458
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
            Index           =   165
            Left            =   1935
            TabIndex        =   457
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   34
         Left            =   -65910
         TabIndex        =   435
         Top             =   5220
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
            Index           =   34
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   442
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
            Index           =   34
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   441
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
            Index           =   34
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   440
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
            Index           =   34
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   439
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
            Index           =   34
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   438
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
            Index           =   34
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   437
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
            Index           =   34
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   436
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
            Index           =   174
            Left            =   90
            TabIndex        =   448
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
            Index           =   173
            Left            =   90
            TabIndex        =   447
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
            Index           =   172
            Left            =   90
            TabIndex        =   446
            Top             =   990
            Width           =   960
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
            TabIndex        =   445
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
            Index           =   171
            Left            =   1035
            TabIndex        =   444
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
            Index           =   170
            Left            =   1935
            TabIndex        =   443
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   35
         Left            =   -62985
         TabIndex        =   421
         Top             =   5220
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
            Index           =   35
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   428
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
            Index           =   35
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   427
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
            Index           =   35
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   426
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
            Index           =   35
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   425
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
            Index           =   35
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   424
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
            Index           =   35
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   423
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
            Index           =   35
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   422
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
            Index           =   179
            Left            =   90
            TabIndex        =   434
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
            Index           =   178
            Left            =   90
            TabIndex        =   433
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
            Index           =   177
            Left            =   90
            TabIndex        =   432
            Top             =   990
            Width           =   960
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
            TabIndex        =   431
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
            Index           =   176
            Left            =   1035
            TabIndex        =   430
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
            Index           =   175
            Left            =   1935
            TabIndex        =   429
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   36
         Left            =   -74685
         TabIndex        =   407
         Top             =   7560
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
            Index           =   36
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   414
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
            Index           =   36
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   413
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
            Index           =   36
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   412
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
            Index           =   36
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   411
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
            Index           =   36
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   410
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
            Index           =   36
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   409
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
            Index           =   36
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   408
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
            Index           =   184
            Left            =   1935
            TabIndex        =   420
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
            Index           =   183
            Left            =   1035
            TabIndex        =   419
            Top             =   675
            Width           =   645
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
            TabIndex        =   418
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
            Index           =   182
            Left            =   90
            TabIndex        =   417
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
            Index           =   181
            Left            =   90
            TabIndex        =   416
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
            Index           =   180
            Left            =   90
            TabIndex        =   415
            Top             =   1800
            Width           =   870
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   37
         Left            =   -71760
         TabIndex        =   393
         Top             =   7560
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
            Index           =   37
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   400
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
            Index           =   37
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   399
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
            Index           =   37
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   398
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
            Index           =   37
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   397
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
            Index           =   37
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   396
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
            Index           =   37
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   395
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
            Index           =   37
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   394
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
            Index           =   189
            Left            =   90
            TabIndex        =   406
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
            Index           =   188
            Left            =   90
            TabIndex        =   405
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
            Index           =   187
            Left            =   90
            TabIndex        =   404
            Top             =   990
            Width           =   960
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
            TabIndex        =   403
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
            Index           =   186
            Left            =   1035
            TabIndex        =   402
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
            Index           =   185
            Left            =   1935
            TabIndex        =   401
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   38
         Left            =   -68835
         TabIndex        =   379
         Top             =   7560
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
            Index           =   38
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   386
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
            Index           =   38
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   385
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
            Index           =   38
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   384
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
            Index           =   38
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   383
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
            Index           =   38
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   382
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
            Index           =   38
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   381
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
            Index           =   38
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   380
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
            Index           =   194
            Left            =   90
            TabIndex        =   392
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
            Index           =   193
            Left            =   90
            TabIndex        =   391
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
            Index           =   192
            Left            =   90
            TabIndex        =   390
            Top             =   990
            Width           =   960
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
            TabIndex        =   389
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
            Index           =   191
            Left            =   1035
            TabIndex        =   388
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
            Index           =   190
            Left            =   1935
            TabIndex        =   387
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   39
         Left            =   -65910
         TabIndex        =   365
         Top             =   7560
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
            Index           =   39
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   372
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
            Index           =   39
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   371
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
            Index           =   39
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   370
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
            Index           =   39
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   369
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
            Index           =   39
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   368
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
            Index           =   39
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   367
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
            Index           =   39
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   366
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
            Index           =   199
            Left            =   90
            TabIndex        =   378
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
            Index           =   198
            Left            =   90
            TabIndex        =   377
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
            Index           =   197
            Left            =   90
            TabIndex        =   376
            Top             =   990
            Width           =   960
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
            TabIndex        =   375
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
            Index           =   196
            Left            =   1035
            TabIndex        =   374
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
            Index           =   195
            Left            =   1935
            TabIndex        =   373
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   40
         Left            =   -62985
         TabIndex        =   351
         Top             =   7560
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
            Index           =   40
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   358
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
            Index           =   40
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   357
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
            Index           =   40
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   356
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
            Index           =   40
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   355
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
            Index           =   40
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   354
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
            Index           =   40
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   353
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
            Index           =   40
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   352
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
            Index           =   204
            Left            =   90
            TabIndex        =   364
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
            Index           =   203
            Left            =   90
            TabIndex        =   363
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
            Index           =   202
            Left            =   90
            TabIndex        =   362
            Top             =   990
            Width           =   960
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
            TabIndex        =   361
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
            Index           =   201
            Left            =   1035
            TabIndex        =   360
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
            Index           =   200
            Left            =   1935
            TabIndex        =   359
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   21
         Left            =   -74685
         TabIndex        =   337
         Top             =   540
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
            TabIndex        =   344
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
            TabIndex        =   343
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
            TabIndex        =   342
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
            TabIndex        =   341
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
            TabIndex        =   340
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
            TabIndex        =   339
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
            TabIndex        =   338
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
            TabIndex        =   350
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
            TabIndex        =   349
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
            TabIndex        =   348
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
            TabIndex        =   347
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
            TabIndex        =   346
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
            TabIndex        =   345
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   22
         Left            =   -71760
         TabIndex        =   323
         Top             =   540
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
            TabIndex        =   330
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
            TabIndex        =   329
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
            TabIndex        =   328
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
            TabIndex        =   327
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
            TabIndex        =   326
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
            TabIndex        =   325
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
            TabIndex        =   324
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
            TabIndex        =   336
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
            TabIndex        =   335
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
            TabIndex        =   334
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
            TabIndex        =   333
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
            TabIndex        =   332
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
            TabIndex        =   331
            Top             =   1800
            Width           =   870
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   23
         Left            =   -68835
         TabIndex        =   309
         Top             =   540
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
            TabIndex        =   316
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
            TabIndex        =   315
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
            TabIndex        =   314
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
            TabIndex        =   313
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
            TabIndex        =   312
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
            TabIndex        =   311
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
            TabIndex        =   310
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
            TabIndex        =   322
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
            TabIndex        =   321
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
            TabIndex        =   320
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
            TabIndex        =   319
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
            TabIndex        =   318
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
            TabIndex        =   317
            Top             =   1800
            Width           =   870
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   24
         Left            =   -65910
         TabIndex        =   295
         Top             =   540
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
            TabIndex        =   302
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
            TabIndex        =   301
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
            TabIndex        =   300
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
            TabIndex        =   299
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
            TabIndex        =   298
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
            TabIndex        =   297
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
            TabIndex        =   296
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
            TabIndex        =   308
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
            TabIndex        =   307
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
            TabIndex        =   306
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
            TabIndex        =   305
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
            TabIndex        =   304
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
            TabIndex        =   303
            Top             =   1800
            Width           =   870
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   25
         Left            =   -62985
         TabIndex        =   281
         Top             =   540
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
            TabIndex        =   288
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
            TabIndex        =   287
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
            TabIndex        =   286
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
            TabIndex        =   285
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
            TabIndex        =   284
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
            TabIndex        =   283
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
            TabIndex        =   282
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
            TabIndex        =   294
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
            TabIndex        =   293
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
            TabIndex        =   292
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
            TabIndex        =   291
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
            TabIndex        =   290
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
            TabIndex        =   289
            Top             =   1800
            Width           =   870
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   0
         Left            =   315
         TabIndex        =   267
         Top             =   540
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
            Index           =   1
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   274
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
            Index           =   1
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   273
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
            Index           =   1
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   272
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
            Index           =   1
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   271
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
            Index           =   1
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   270
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
            Index           =   1
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   269
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
            Index           =   1
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   268
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
            Index           =   4
            Left            =   1935
            TabIndex        =   280
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
            Index           =   3
            Left            =   1035
            TabIndex        =   279
            Top             =   675
            Width           =   645
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
            TabIndex        =   278
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
            Index           =   0
            Left            =   90
            TabIndex        =   277
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
            Index           =   1
            Left            =   90
            TabIndex        =   276
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
            Index           =   2
            Left            =   90
            TabIndex        =   275
            Top             =   1800
            Width           =   870
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   2
         Left            =   3240
         TabIndex        =   253
         Top             =   540
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
            Index           =   2
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   260
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
            Index           =   2
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   259
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
            Index           =   2
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   258
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
            Index           =   2
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   257
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
            Index           =   2
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   256
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
            Index           =   2
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   255
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
            Index           =   2
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   254
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
            Index           =   14
            Left            =   90
            TabIndex        =   266
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
            Index           =   13
            Left            =   90
            TabIndex        =   265
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
            Index           =   12
            Left            =   90
            TabIndex        =   264
            Top             =   990
            Width           =   960
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
            TabIndex        =   263
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
            Index           =   11
            Left            =   1035
            TabIndex        =   262
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
            Index           =   10
            Left            =   1935
            TabIndex        =   261
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   3
         Left            =   6165
         TabIndex        =   239
         Top             =   540
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
            Index           =   3
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   246
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
            Index           =   3
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   245
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
            Index           =   3
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   244
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
            Index           =   3
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   243
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
            Index           =   3
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   242
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
            Index           =   3
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   241
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
            Index           =   3
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   240
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
            Index           =   19
            Left            =   90
            TabIndex        =   252
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
            Index           =   18
            Left            =   90
            TabIndex        =   251
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
            Index           =   17
            Left            =   90
            TabIndex        =   250
            Top             =   990
            Width           =   960
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
            TabIndex        =   249
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
            Index           =   16
            Left            =   1035
            TabIndex        =   248
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
            Index           =   15
            Left            =   1935
            TabIndex        =   247
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   4
         Left            =   9090
         TabIndex        =   225
         Top             =   540
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
            Index           =   4
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   232
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
            Index           =   4
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   231
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
            Index           =   4
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   230
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
            Index           =   4
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   229
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
            Index           =   4
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   228
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
            Index           =   4
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   227
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
            Index           =   4
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   226
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
            Index           =   24
            Left            =   90
            TabIndex        =   238
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
            Index           =   23
            Left            =   90
            TabIndex        =   237
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
            Index           =   22
            Left            =   90
            TabIndex        =   236
            Top             =   990
            Width           =   960
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
            TabIndex        =   235
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
            Index           =   21
            Left            =   1035
            TabIndex        =   234
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
            Index           =   20
            Left            =   1935
            TabIndex        =   233
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   5
         Left            =   12015
         TabIndex        =   211
         Top             =   540
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
            Index           =   5
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   218
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
            Index           =   5
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   217
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
            Index           =   5
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   216
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
            Index           =   5
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   215
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
            Index           =   5
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   214
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
            Index           =   5
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   213
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
            Index           =   5
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   212
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
            Index           =   29
            Left            =   90
            TabIndex        =   224
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
            Index           =   28
            Left            =   90
            TabIndex        =   223
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
            Index           =   27
            Left            =   90
            TabIndex        =   222
            Top             =   990
            Width           =   960
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
            TabIndex        =   221
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
            Index           =   26
            Left            =   1035
            TabIndex        =   220
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
            Index           =   25
            Left            =   1935
            TabIndex        =   219
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   6
         Left            =   315
         TabIndex        =   197
         Top             =   2880
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
            Index           =   6
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   204
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
            Index           =   6
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   203
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
            Index           =   6
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   202
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
            Index           =   6
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   201
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
            Index           =   6
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   200
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
            Index           =   6
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   199
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
            Index           =   6
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   198
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
            Index           =   34
            Left            =   90
            TabIndex        =   210
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
            Index           =   33
            Left            =   90
            TabIndex        =   209
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
            Index           =   32
            Left            =   90
            TabIndex        =   208
            Top             =   990
            Width           =   960
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
            TabIndex        =   207
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
            Index           =   31
            Left            =   1035
            TabIndex        =   206
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
            Index           =   30
            Left            =   1935
            TabIndex        =   205
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   7
         Left            =   3240
         TabIndex        =   183
         Top             =   2880
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
            Index           =   7
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   190
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
            Index           =   7
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   189
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
            Index           =   7
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   188
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
            Index           =   7
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   187
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
            Index           =   7
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   186
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
            Index           =   7
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   185
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
            Index           =   7
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   184
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
            Index           =   39
            Left            =   1935
            TabIndex        =   196
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
            Index           =   38
            Left            =   1035
            TabIndex        =   195
            Top             =   675
            Width           =   645
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
            TabIndex        =   194
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
            Index           =   37
            Left            =   90
            TabIndex        =   193
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
            Index           =   36
            Left            =   90
            TabIndex        =   192
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
            Index           =   35
            Left            =   90
            TabIndex        =   191
            Top             =   1800
            Width           =   870
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   8
         Left            =   6165
         TabIndex        =   169
         Top             =   2880
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
            Index           =   8
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   176
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
            Index           =   8
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   175
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
            Index           =   8
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   174
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
            Index           =   8
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   173
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
            Index           =   8
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   172
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
            Index           =   8
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   171
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
            Index           =   8
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   170
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
            Index           =   44
            Left            =   1935
            TabIndex        =   182
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
            Index           =   43
            Left            =   1035
            TabIndex        =   181
            Top             =   675
            Width           =   645
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
            TabIndex        =   180
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
            Index           =   42
            Left            =   90
            TabIndex        =   179
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
            Index           =   41
            Left            =   90
            TabIndex        =   178
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
            Index           =   40
            Left            =   90
            TabIndex        =   177
            Top             =   1800
            Width           =   870
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   9
         Left            =   9090
         TabIndex        =   155
         Top             =   2880
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
            Index           =   9
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   162
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
            Index           =   9
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   161
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
            Index           =   9
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   160
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
            Index           =   9
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   159
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
            Index           =   9
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   158
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
            Index           =   9
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   157
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
            Index           =   9
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   156
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
            Index           =   49
            Left            =   1935
            TabIndex        =   168
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
            Index           =   48
            Left            =   1035
            TabIndex        =   167
            Top             =   675
            Width           =   645
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
            TabIndex        =   166
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
            Index           =   47
            Left            =   90
            TabIndex        =   165
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
            Index           =   46
            Left            =   90
            TabIndex        =   164
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
            Index           =   45
            Left            =   90
            TabIndex        =   163
            Top             =   1800
            Width           =   870
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   10
         Left            =   12015
         TabIndex        =   141
         Top             =   2880
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
            Index           =   10
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   148
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
            Index           =   10
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   147
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
            Index           =   10
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   146
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
            Index           =   10
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   145
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
            Index           =   10
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   144
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
            Index           =   10
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   143
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
            Index           =   10
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   142
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
            Index           =   54
            Left            =   1935
            TabIndex        =   154
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
            Index           =   53
            Left            =   1035
            TabIndex        =   153
            Top             =   675
            Width           =   645
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
            TabIndex        =   152
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
            Index           =   52
            Left            =   90
            TabIndex        =   151
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
            Index           =   51
            Left            =   90
            TabIndex        =   150
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
            Index           =   50
            Left            =   90
            TabIndex        =   149
            Top             =   1800
            Width           =   870
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   11
         Left            =   315
         TabIndex        =   127
         Top             =   5220
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
            Index           =   11
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   134
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
            Index           =   11
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   133
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
            Index           =   11
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   132
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
            Index           =   11
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   131
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
            Index           =   11
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   130
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
            Index           =   11
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   129
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
            Index           =   11
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   128
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
            Index           =   59
            Left            =   90
            TabIndex        =   140
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
            Index           =   58
            Left            =   90
            TabIndex        =   139
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
            Index           =   57
            Left            =   90
            TabIndex        =   138
            Top             =   990
            Width           =   960
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
            TabIndex        =   137
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
            Index           =   56
            Left            =   1035
            TabIndex        =   136
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
            Index           =   55
            Left            =   1935
            TabIndex        =   135
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   12
         Left            =   3240
         TabIndex        =   113
         Top             =   5220
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
            Index           =   12
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   120
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
            Index           =   12
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   119
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
            Index           =   12
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   118
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
            Index           =   12
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   117
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
            Index           =   12
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   116
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
            Index           =   12
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   115
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
            Index           =   12
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   114
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
            Index           =   64
            Left            =   1935
            TabIndex        =   126
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
            Index           =   63
            Left            =   1035
            TabIndex        =   125
            Top             =   675
            Width           =   645
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
            TabIndex        =   124
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
            Index           =   62
            Left            =   90
            TabIndex        =   123
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
            Index           =   61
            Left            =   90
            TabIndex        =   122
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
            Index           =   60
            Left            =   90
            TabIndex        =   121
            Top             =   1800
            Width           =   870
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   13
         Left            =   6165
         TabIndex        =   99
         Top             =   5220
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
            Index           =   13
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   106
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
            Index           =   13
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   105
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
            Index           =   13
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   104
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
            Index           =   13
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   103
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
            Index           =   13
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   102
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
            Index           =   13
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   101
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
            Index           =   13
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   100
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
            Index           =   69
            Left            =   1935
            TabIndex        =   112
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
            Index           =   68
            Left            =   1035
            TabIndex        =   111
            Top             =   675
            Width           =   645
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
            TabIndex        =   110
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
            Index           =   67
            Left            =   90
            TabIndex        =   109
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
            Index           =   66
            Left            =   90
            TabIndex        =   108
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
            Index           =   65
            Left            =   90
            TabIndex        =   107
            Top             =   1800
            Width           =   870
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   14
         Left            =   9090
         TabIndex        =   85
         Top             =   5220
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
            Index           =   14
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   92
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
            Index           =   14
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   91
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
            Index           =   14
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   90
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
            Index           =   14
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   89
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
            Index           =   14
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   88
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
            Index           =   14
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   87
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
            Index           =   14
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   86
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
            Index           =   74
            Left            =   1935
            TabIndex        =   98
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
            Index           =   73
            Left            =   1035
            TabIndex        =   97
            Top             =   675
            Width           =   645
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
            TabIndex        =   96
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
            Index           =   72
            Left            =   90
            TabIndex        =   95
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
            Index           =   71
            Left            =   90
            TabIndex        =   94
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
            Index           =   70
            Left            =   90
            TabIndex        =   93
            Top             =   1800
            Width           =   870
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   15
         Left            =   12015
         TabIndex        =   71
         Top             =   5220
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
            Index           =   15
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   78
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
            Index           =   15
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   77
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
            Index           =   15
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   76
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
            Index           =   15
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   75
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
            Index           =   15
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   74
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
            Index           =   15
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   73
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
            Index           =   15
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   72
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
            Index           =   79
            Left            =   1935
            TabIndex        =   84
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
            Index           =   78
            Left            =   1035
            TabIndex        =   83
            Top             =   675
            Width           =   645
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
            TabIndex        =   82
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
            Index           =   77
            Left            =   90
            TabIndex        =   81
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
            Index           =   76
            Left            =   90
            TabIndex        =   80
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
            Index           =   75
            Left            =   90
            TabIndex        =   79
            Top             =   1800
            Width           =   870
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   16
         Left            =   315
         TabIndex        =   57
         Top             =   7560
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
            Index           =   16
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   64
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
            Index           =   16
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   63
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
            Index           =   16
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   62
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
            Index           =   16
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   61
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
            Index           =   16
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   60
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
            Index           =   16
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   59
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
            Index           =   16
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   58
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
            Index           =   84
            Left            =   1935
            TabIndex        =   70
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
            Index           =   83
            Left            =   1035
            TabIndex        =   69
            Top             =   675
            Width           =   645
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
            TabIndex        =   68
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
            Index           =   82
            Left            =   90
            TabIndex        =   67
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
            Index           =   81
            Left            =   90
            TabIndex        =   66
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
            Index           =   80
            Left            =   90
            TabIndex        =   65
            Top             =   1800
            Width           =   870
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   17
         Left            =   3240
         TabIndex        =   43
         Top             =   7560
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
            Index           =   17
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   50
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
            Index           =   17
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   49
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
            Index           =   17
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   48
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
            Index           =   17
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   47
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
            Index           =   17
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   46
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
            Index           =   17
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   45
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
            Index           =   17
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   44
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
            Index           =   89
            Left            =   90
            TabIndex        =   56
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
            Index           =   88
            Left            =   90
            TabIndex        =   55
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
            Index           =   87
            Left            =   90
            TabIndex        =   54
            Top             =   990
            Width           =   960
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
            TabIndex        =   53
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
            Index           =   86
            Left            =   1035
            TabIndex        =   52
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
            Index           =   85
            Left            =   1935
            TabIndex        =   51
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   18
         Left            =   6165
         TabIndex        =   29
         Top             =   7560
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
            Index           =   18
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   36
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
            Index           =   18
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   35
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
            Index           =   18
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   34
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
            Index           =   18
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   33
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
            Index           =   18
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   32
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
            Index           =   18
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   31
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
            Index           =   18
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   30
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
            Index           =   94
            Left            =   90
            TabIndex        =   42
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
            Index           =   93
            Left            =   90
            TabIndex        =   41
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
            Index           =   92
            Left            =   90
            TabIndex        =   40
            Top             =   990
            Width           =   960
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
            TabIndex        =   39
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
            Index           =   91
            Left            =   1035
            TabIndex        =   38
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
            Index           =   90
            Left            =   1935
            TabIndex        =   37
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   19
         Left            =   9090
         TabIndex        =   15
         Top             =   7560
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
            Index           =   19
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   22
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
            Index           =   19
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   21
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
            Index           =   19
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   20
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
            Index           =   19
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   19
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
            Index           =   19
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   18
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
            Index           =   19
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   17
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
            Index           =   19
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   16
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
            Index           =   99
            Left            =   90
            TabIndex        =   28
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
            Index           =   98
            Left            =   90
            TabIndex        =   27
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
            Index           =   97
            Left            =   90
            TabIndex        =   26
            Top             =   990
            Width           =   960
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
            TabIndex        =   25
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
            Index           =   96
            Left            =   1035
            TabIndex        =   24
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
            Index           =   95
            Left            =   1935
            TabIndex        =   23
            Top             =   675
            Width           =   645
         End
      End
      Begin VB.Frame frameDesign 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Index           =   20
         Left            =   12015
         TabIndex        =   1
         Top             =   7560
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
            Index           =   20
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   8
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
            Index           =   20
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   7
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
            Index           =   20
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   6
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
            Index           =   20
            Left            =   585
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   5
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
            Index           =   20
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   4
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
            Index           =   20
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   3
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
            Index           =   20
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   2
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
            Index           =   104
            Left            =   90
            TabIndex        =   14
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
            Index           =   103
            Left            =   90
            TabIndex        =   13
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
            Index           =   102
            Left            =   90
            TabIndex        =   12
            Top             =   990
            Width           =   960
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
            TabIndex        =   11
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
            Index           =   101
            Left            =   1035
            TabIndex        =   10
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
            Index           =   100
            Left            =   1935
            TabIndex        =   9
            Top             =   675
            Width           =   645
         End
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   135
      TabIndex        =   592
      Top             =   11250
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   405
      Left            =   13995
      TabIndex        =   604
      Top             =   90
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
      Caption         =   "&Close [ESC]"
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
      cFHover         =   4210752
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmYieldReport.frx":6852
      cBack           =   16119285
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
      Left            =   45
      TabIndex        =   593
      Top             =   45
      Width           =   15495
   End
End
Attribute VB_Name = "frmYieldReport"
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



Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub DTDate_Change()
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
'                tBox.BackColor = &HC000C0
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


Private Sub Form_Activate()
On Error Resume Next
With MAIN
    Me.BackColor = .ACPMenu.BackColor
    Me.Picture = .ACPMenu.LoadBackground
End With

MAIN.ActivateChild Me


End Sub
