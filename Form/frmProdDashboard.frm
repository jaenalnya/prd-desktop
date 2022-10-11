VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{8E048CF2-F435-45C9-8A6F-4646F9E1B5F4}#1.0#0"; "XPTab.ocx"
Begin VB.Form frmProdDashboard 
   Caption         =   "Productivity Dashboard"
   ClientHeight    =   10245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16680
   Icon            =   "frmProdDashboard.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10245
   ScaleWidth      =   16680
   Begin VB.Frame frame 
      BackColor       =   &H00404040&
      Height          =   1185
      Index           =   15
      Left            =   13815
      TabIndex        =   69
      Top             =   8190
      Width           =   2760
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Good Units"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   135
         TabIndex        =   71
         Top             =   180
         Width           =   2400
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   330
         Left            =   225
         TabIndex        =   70
         Top             =   585
         Width           =   2310
      End
   End
   Begin VB.Frame frame 
      BackColor       =   &H00404040&
      Height          =   1185
      Index           =   14
      Left            =   11070
      TabIndex        =   66
      Top             =   8190
      Width           =   2715
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   225
         TabIndex        =   68
         Top             =   585
         Width           =   2310
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rejected Units"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   135
         TabIndex        =   67
         Top             =   180
         Width           =   2400
      End
   End
   Begin VB.Frame frame 
      BackColor       =   &H00404040&
      Height          =   1185
      Index           =   13
      Left            =   8325
      TabIndex        =   63
      Top             =   8190
      Width           =   2715
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Produced Units"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   135
         TabIndex        =   65
         Top             =   180
         Width           =   2400
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   330
         Left            =   225
         TabIndex        =   64
         Top             =   585
         Width           =   2310
      End
   End
   Begin VB.Frame frame 
      BackColor       =   &H00404040&
      Height          =   1185
      Index           =   12
      Left            =   13815
      TabIndex        =   60
      Top             =   6975
      Width           =   2760
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   330
         Left            =   225
         TabIndex        =   62
         Top             =   585
         Width           =   2310
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Good Units"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   135
         TabIndex        =   61
         Top             =   180
         Width           =   2400
      End
   End
   Begin VB.Frame frame 
      BackColor       =   &H00404040&
      Height          =   1185
      Index           =   9
      Left            =   11070
      TabIndex        =   57
      Top             =   6975
      Width           =   2715
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rejected Units"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   135
         TabIndex        =   59
         Top             =   180
         Width           =   2400
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   225
         TabIndex        =   58
         Top             =   585
         Width           =   2310
      End
   End
   Begin VB.Frame frame 
      BackColor       =   &H00404040&
      Height          =   555
      Index           =   8
      Left            =   8325
      TabIndex        =   45
      Top             =   1350
      Width           =   8250
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "120"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   2700
         TabIndex        =   48
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Actual Production :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   180
         TabIndex        =   46
         Top             =   180
         Width           =   2355
      End
   End
   Begin VB.Frame frame 
      BackColor       =   &H00404040&
      Height          =   555
      Index           =   6
      Left            =   8325
      TabIndex        =   42
      Top             =   765
      Width           =   8250
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "(unit / hour)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   5085
         TabIndex        =   47
         Top             =   180
         Width           =   3840
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Production Capacity :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   180
         TabIndex        =   44
         Top             =   180
         Width           =   2580
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "120"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2700
         TabIndex        =   43
         Top             =   180
         Width           =   915
      End
   End
   Begin VB.Frame frame 
      BackColor       =   &H00404040&
      Height          =   555
      Index           =   5
      Left            =   45
      TabIndex        =   28
      Top             =   1350
      Width           =   8250
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   180
         TabIndex        =   30
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "PT. Furukawa Indomobil Battery Manufacturing"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1575
         TabIndex        =   29
         Top             =   180
         Width           =   6540
      End
   End
   Begin VB.Frame frame 
      BackColor       =   &H00404040&
      Height          =   825
      Index           =   17
      Left            =   45
      TabIndex        =   8
      Top             =   9360
      Width           =   16575
   End
   Begin VB.Frame frame 
      BackColor       =   &H00404040&
      Height          =   1185
      Index           =   11
      Left            =   8325
      TabIndex        =   7
      Top             =   6975
      Width           =   2715
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   330
         Left            =   225
         TabIndex        =   56
         Top             =   585
         Width           =   2310
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Produced Units"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   135
         TabIndex        =   55
         Top             =   180
         Width           =   2400
      End
   End
   Begin VB.Frame frame 
      BackColor       =   &H00404040&
      ForeColor       =   &H00404040&
      Height          =   690
      Index           =   0
      Left            =   45
      TabIndex        =   6
      Top             =   45
      Width           =   16530
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "2020-03-01  11:30:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   14400
         TabIndex        =   15
         Top             =   270
         Width           =   2130
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   9720
         TabIndex        =   14
         Top             =   270
         Width           =   645
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Active Alarms :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   8370
         TabIndex        =   13
         Top             =   270
         Width           =   1410
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "JAENAL NUROHMAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5805
         TabIndex        =   12
         Top             =   270
         Width           =   2490
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "31811914"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   4860
         TabIndex        =   11
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Active User :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3600
         TabIndex        =   10
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Productivity Dashboard"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   180
         TabIndex        =   9
         Top             =   180
         Width           =   3390
      End
   End
   Begin VB.Frame frame 
      BackColor       =   &H00404040&
      Height          =   5550
      Index           =   7
      Left            =   45
      TabIndex        =   5
      Top             =   3825
      Width           =   8250
      Begin PRD.Liner Liner1 
         Height          =   30
         Left            =   135
         TabIndex        =   37
         Top             =   2835
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   53
      End
      Begin prjXTab.XTab XTab1 
         Height          =   2220
         Left            =   180
         TabIndex        =   31
         Top             =   450
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   3916
         TabCaption(0)   =   "Shift 1"
         TabContCtrlCnt(0)=   1
         Tab(0)ContCtrlCap(1)=   "lvList"
         TabCaption(1)   =   "Shift 2"
         TabContCtrlCnt(1)=   1
         Tab(1)ContCtrlCap(1)=   "ListView2"
         TabCaption(2)   =   "Shift 3"
         TabContCtrlCnt(2)=   1
         Tab(2)ContCtrlCap(1)=   "ListView3"
         TabTheme        =   1
         ActiveTabBackStartColor=   16514555
         ActiveTabBackEndColor=   16514555
         InActiveTabBackStartColor=   16777215
         InActiveTabBackEndColor=   15397104
         BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OuterBorderColor=   10198161
         DisabledTabBackColor=   -2147483633
         DisabledTabForeColor=   10526880
         Begin MSComctlLib.ListView ListView3 
            Height          =   1770
            Left            =   -74955
            TabIndex        =   39
            Top             =   405
            Width           =   7800
            _ExtentX        =   13758
            _ExtentY        =   3122
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "i16x16"
            SmallIcons      =   "i16x16"
            ColHdrIcons     =   "i16x16"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   1770
            Left            =   -74955
            TabIndex        =   38
            Top             =   405
            Width           =   7800
            _ExtentX        =   13758
            _ExtentY        =   3122
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "i16x16"
            SmallIcons      =   "i16x16"
            ColHdrIcons     =   "i16x16"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin MSComctlLib.ListView lvList 
            Height          =   1770
            Left            =   45
            TabIndex        =   33
            Top             =   405
            Width           =   7800
            _ExtentX        =   13758
            _ExtentY        =   3122
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "i16x16"
            SmallIcons      =   "i16x16"
            ColHdrIcons     =   "i16x16"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin prjXTab.XTab XTab2 
         Height          =   2220
         Left            =   180
         TabIndex        =   32
         Top             =   3195
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   3916
         TabCaption(0)   =   "Shift 1"
         TabContCtrlCnt(0)=   1
         Tab(0)ContCtrlCap(1)=   "ListView1"
         TabCaption(1)   =   "Shift 2"
         TabContCtrlCnt(1)=   1
         Tab(1)ContCtrlCap(1)=   "ListView4"
         TabCaption(2)   =   "Shift 3"
         TabContCtrlCnt(2)=   1
         Tab(2)ContCtrlCap(1)=   "ListView5"
         TabTheme        =   1
         ActiveTabBackStartColor=   16514555
         ActiveTabBackEndColor=   16514555
         InActiveTabBackStartColor=   16777215
         InActiveTabBackEndColor=   15397104
         BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OuterBorderColor=   10198161
         DisabledTabBackColor=   -2147483633
         DisabledTabForeColor=   10526880
         Begin MSComctlLib.ListView ListView5 
            Height          =   1770
            Left            =   -74955
            TabIndex        =   41
            Top             =   405
            Width           =   7800
            _ExtentX        =   13758
            _ExtentY        =   3122
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "i16x16"
            SmallIcons      =   "i16x16"
            ColHdrIcons     =   "i16x16"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin MSComctlLib.ListView ListView4 
            Height          =   1770
            Left            =   -74955
            TabIndex        =   40
            Top             =   405
            Width           =   7800
            _ExtentX        =   13758
            _ExtentY        =   3122
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "i16x16"
            SmallIcons      =   "i16x16"
            ColHdrIcons     =   "i16x16"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1770
            Left            =   45
            TabIndex        =   34
            Top             =   405
            Width           =   7800
            _ExtentX        =   13758
            _ExtentY        =   3122
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "i16x16"
            SmallIcons      =   "i16x16"
            ColHdrIcons     =   "i16x16"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCT 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   36
         Top             =   2925
         Width           =   1140
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCT 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   35
         Top             =   180
         Width           =   1140
      End
   End
   Begin VB.Frame frame 
      BackColor       =   &H00404040&
      Height          =   915
      Index           =   3
      Left            =   45
      TabIndex        =   4
      Top             =   2880
      Width           =   4110
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1620
         TabIndex        =   27
         Top             =   540
         Width           =   6585
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Shift :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   225
         TabIndex        =   26
         Top             =   540
         Width           =   1185
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "2020-03-01  11:30:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1620
         TabIndex        =   23
         Top             =   225
         Width           =   6585
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Date :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   225
         TabIndex        =   22
         Top             =   225
         Width           =   1185
      End
   End
   Begin VB.Frame frame 
      BackColor       =   &H00404040&
      Height          =   5010
      Index           =   10
      Left            =   8325
      TabIndex        =   3
      Top             =   1935
      Width           =   8250
      Begin prjXTab.XTab XTab3 
         Height          =   4110
         Left            =   180
         TabIndex        =   49
         Top             =   585
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   7250
         TabCount        =   2
         TabCaption(0)   =   "Data NG"
         TabContCtrlCnt(0)=   1
         Tab(0)ContCtrlCap(1)=   "ListView9"
         TabCaption(1)   =   "Data Idle Time"
         TabContCtrlCnt(1)=   1
         Tab(1)ContCtrlCap(1)=   "ListView7"
         TabTheme        =   1
         ActiveTabBackStartColor=   16514555
         ActiveTabBackEndColor=   16514555
         InActiveTabBackStartColor=   16777215
         InActiveTabBackEndColor=   15397104
         BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OuterBorderColor=   10198161
         DisabledTabBackColor=   -2147483633
         DisabledTabForeColor=   10526880
         Begin MSComctlLib.ListView ListView9 
            Height          =   3525
            Left            =   45
            TabIndex        =   53
            Top             =   495
            Width           =   7800
            _ExtentX        =   13758
            _ExtentY        =   6218
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "i16x16"
            SmallIcons      =   "i16x16"
            ColHdrIcons     =   "i16x16"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin MSComctlLib.ListView ListView8 
            Height          =   1770
            Left            =   -74955
            TabIndex        =   52
            Top             =   405
            Width           =   7800
            _ExtentX        =   13758
            _ExtentY        =   3122
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "i16x16"
            SmallIcons      =   "i16x16"
            ColHdrIcons     =   "i16x16"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin MSComctlLib.ListView ListView7 
            Height          =   3570
            Left            =   -74955
            TabIndex        =   51
            Top             =   405
            Width           =   7800
            _ExtentX        =   13758
            _ExtentY        =   6297
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "i16x16"
            SmallIcons      =   "i16x16"
            ColHdrIcons     =   "i16x16"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin MSComctlLib.ListView ListView6 
            Height          =   1770
            Left            =   -74955
            TabIndex        =   50
            Top             =   405
            Width           =   7800
            _ExtentX        =   13758
            _ExtentY        =   3122
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "i16x16"
            SmallIcons      =   "i16x16"
            ColHdrIcons     =   "i16x16"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TOP 10 DATA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   135
         TabIndex        =   54
         Top             =   225
         Width           =   7890
      End
   End
   Begin VB.Frame frame 
      BackColor       =   &H00404040&
      Height          =   915
      Index           =   4
      Left            =   4185
      TabIndex        =   2
      Top             =   2880
      Width           =   4110
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "RUNNING"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   1620
         TabIndex        =   25
         Top             =   225
         Width           =   6585
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Status :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   225
         TabIndex        =   24
         Top             =   225
         Width           =   1185
      End
   End
   Begin VB.Frame frame 
      BackColor       =   &H00404040&
      Height          =   555
      Index           =   1
      Left            =   45
      TabIndex        =   1
      Top             =   765
      Width           =   8250
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Toshiba 280"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   2295
         TabIndex        =   18
         Top             =   180
         Width           =   1995
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "41"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   1620
         TabIndex        =   17
         Top             =   180
         Width           =   645
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Machine :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   180
         TabIndex        =   16
         Top             =   180
         Width           =   1230
      End
   End
   Begin VB.Frame frame 
      BackColor       =   &H00404040&
      Height          =   915
      Index           =   2
      Left            =   45
      TabIndex        =   0
      Top             =   1935
      Width           =   8250
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "PIB003R00060000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1620
         TabIndex        =   21
         Top             =   225
         Width           =   6540
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Product :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   225
         TabIndex        =   20
         Top             =   225
         Width           =   1185
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "R-DOOR CAP TOP SILVER GREY (CROWN SERIES) 1 DOOR PRINT (503) WIP	"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1620
         TabIndex        =   19
         Top             =   540
         Width           =   6540
      End
   End
End
Attribute VB_Name = "frmProdDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
On Error Resume Next
With MAIN
    Me.BackColor = .ACPMenu.BackColor
    Me.Picture = .ACPMenu.LoadBackground
    Dim i As Integer
    For i = 0 To 17
        frame(i).BackColor = .ACPMenu.BackColor
    Next i
End With

End Sub
