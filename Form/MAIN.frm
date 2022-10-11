VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{248850FC-2BAF-48AF-99D6-220E54FE68CA}#1.0#0"; "HookMenu.ocx"
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "B8CONT~2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MAIN 
   BackColor       =   &H8000000C&
   Caption         =   "Production Monitoring System"
   ClientHeight    =   12420
   ClientLeft      =   2625
   ClientTop       =   1710
   ClientWidth     =   21765
   Icon            =   "MAIN.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   21765
      _ExtentX        =   38391
      _ExtentY        =   1535
      ButtonWidth     =   1984
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "DASHBOARD"
            Key             =   "dashboard"
            Object.ToolTipText     =   "Show Dashboard"
            ImageIndex      =   69
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " WI/CCP/SP "
            Key             =   "WI/CCP/SP"
            Object.ToolTipText     =   "Show WI/CCP/SP"
            ImageIndex      =   70
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "INFO"
            Key             =   "Info"
            Object.ToolTipText     =   "Show Info"
            ImageIndex      =   99
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "CALL MTC"
            Key             =   "Call/SMS"
            Object.ToolTipText     =   "Show Call/SMS"
            ImageIndex      =   94
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "PROD RESULT"
            Key             =   "Prod Result"
            Object.ToolTipText     =   "Show Prod Result"
            ImageIndex      =   82
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "NG / REJECT"
            Key             =   "NG/Reject"
            Object.ToolTipText     =   "Show NG/Reject"
            ImageIndex      =   32
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "IDLE TIME"
            Key             =   "Idle Time"
            Object.ToolTipText     =   "Show Idle Time"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "TROUBLE"
            Key             =   "Trouble Machine"
            Object.ToolTipText     =   "Trouble Machine"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "MONITORING"
            Key             =   "Monitoring Leader"
            Object.ToolTipText     =   "Monitoring Leader"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "CHANGE USR"
            Key             =   "Change User"
            Object.ToolTipText     =   "Change User"
            ImageIndex      =   58
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "LOGIN USR 2"
            Key             =   "Login User 2"
            Object.ToolTipText     =   "Login User 2"
            ImageIndex      =   93
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.Frame frameSensor 
         Height          =   825
         Left            =   14625
         TabIndex        =   24
         Top             =   -45
         Width           =   4200
         Begin VB.Label lblCT 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3195
            TabIndex        =   32
            Top             =   270
            Width           =   870
         End
         Begin VB.Shape ShapeKeyboard 
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   555
            Left            =   1890
            Top             =   180
            Width           =   555
         End
         Begin VB.Label lbloff 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "off"
            Height          =   240
            Left            =   2520
            TabIndex        =   30
            Top             =   450
            Width           =   510
         End
         Begin VB.Label lblOn 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "on"
            Height          =   240
            Left            =   2520
            TabIndex        =   29
            Top             =   225
            Width           =   510
         End
         Begin VB.Label Label1 
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
            Height          =   285
            Left            =   90
            TabIndex        =   28
            Top             =   450
            Width           =   870
         End
         Begin VB.Label Label2 
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
            Height          =   285
            Left            =   990
            TabIndex        =   27
            Top             =   450
            Width           =   870
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   26
            Top             =   180
            Width           =   870
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   990
            TabIndex        =   25
            Top             =   180
            Width           =   870
         End
         Begin VB.Shape Shape1 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H0000FFFF&
            FillStyle       =   0  'Solid
            Height          =   555
            Left            =   2475
            Top             =   180
            Width           =   645
         End
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2475
      Top             =   10890
   End
   Begin VB.Timer TmrInformasi 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   270
      Top             =   10980
   End
   Begin MSComDlg.CommonDialog CDExporter 
      Left            =   12690
      Top             =   10710
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1935
      Top             =   10935
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3960
      Top             =   1755
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   123
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":169B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1CB3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":22CC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":28E50
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2EFDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":35164
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":3B2EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":41478
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":47602
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":4D78C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":53916
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":59AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":5FC2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":65DB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":6BF3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":720C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":78252
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":7E3DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":84566
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":8A6F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":9087A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":96A04
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":9CB8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":A2D18
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":A8EA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":AF02C
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":B51B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":BB340
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":C14CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":C7654
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":CD7DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":D3968
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":D9AF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":DFC7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":E5E06
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":EBF90
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":F211A
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":F82A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":FE42E
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1045B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":10A742
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1108CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":116A56
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":11CBE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":122D6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":128EF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":12F07E
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":135208
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":13B392
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":14151C
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1476A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":14D830
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1539BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":159B44
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":15FCCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":165E58
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":16BFE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":17216C
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1782F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":17E480
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":18460A
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":18A794
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":19091E
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":196AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":19CC32
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1A2DBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1A8F46
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1AF0D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1B525A
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1BB3E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1C156E
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1C76F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1CD882
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1D3A0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1D9B96
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1DFD20
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1E5EAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1EC034
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1F21BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1F8348
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":1FE4D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":20465C
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":20A7E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":210970
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":216AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":21CC84
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":222E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":228F98
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":22F122
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2352AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":23B436
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2415C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":24774A
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":24D8D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":253A5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":259BE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":25FD72
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":265EFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":26C086
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":272210
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":27839A
            Key             =   ""
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":27E524
            Key             =   ""
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2846AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":28A838
            Key             =   ""
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2909C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":296B4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":29CCD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2A2E60
            Key             =   ""
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2A8FEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2AF174
            Key             =   ""
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2B52FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2BB488
            Key             =   ""
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2C1612
            Key             =   ""
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2C779C
            Key             =   ""
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2CD926
            Key             =   ""
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2D3AB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2D9C3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2DFDC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2E5F4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2EC0D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2F2262
            Key             =   ""
         EndProperty
         BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2F83EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":2FE576
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   810
      Top             =   10980
   End
   Begin VB.Timer TimerPort 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1350
      Top             =   10980
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   10
      Index           =   5
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   21765
      TabIndex        =   2
      Top             =   12060
      Width           =   21765
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   10
      Index           =   4
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   21765
      TabIndex        =   1
      Top             =   12045
      Width           =   21765
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   2
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   21765
      TabIndex        =   0
      Top             =   12030
      Width           =   21765
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   12075
      Width           =   21765
      _ExtentX        =   38391
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   15
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2293
            MinWidth        =   2293
            Text            =   "ACTIVE USER 1:"
            TextSave        =   "ACTIVE USER 1:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8334
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2293
            MinWidth        =   2293
            Text            =   "ACTIVE USER 2:"
            TextSave        =   "ACTIVE USER 2:"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8334
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Today:"
            TextSave        =   "Today:"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "8/24/2022"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "17:45"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel13 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel14 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel15 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin HookMenu.ctxHookMenu HookMenu 
      Left            =   11340
      Top             =   9945
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   9
      Key:1           =   "#mnuRACN"
      Mask:2          =   16777215
      Key:2           =   "#mnuRAES"
      Key:3           =   "#mnuRAP"
      Mask:4          =   16777215
      Key:4           =   "#mnuRADS"
      Key:5           =   "#mnuRARR"
      Mask:6          =   16777215
      Key:6           =   "#mnuRAC"
      Key:7           =   "#mUCM"
      Key:8           =   "#mCM"
      Mask:9          =   16777215
      Key:9           =   "#mnuSRC"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11610
      Top             =   8910
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   90
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":304700
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":306092
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":306D6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":308700
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":30A092
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":30BA24
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":30D3B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":30D9CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":30E6A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":30F380
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":31005A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":310D36
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":311A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":3122EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":312583
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":31325F
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":313B43
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":31407E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":31479C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":314996
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":315270
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":3158E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":315BCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":315FB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":3169AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":316E5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":317072
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":3176CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":317C5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":31822B
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":318B05
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":3193DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":330479
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":347513
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":35E5AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":375647
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":38C6E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":3A377B
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":3BA815
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":3D18AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":3E8949
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":3FF9E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":416A7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":42DB17
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":444BB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":45BF43
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":472FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":48A077
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":4A1111
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":4B81AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":4CF245
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":4E62DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":4FD379
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":514413
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":52B4AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":542547
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":5595E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":57067B
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":587715
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":59E7AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":5B5849
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":5CC8E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":5E397D
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":5FAA17
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":611AB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":628B4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":63FBE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":656C7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":66DD19
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":684DB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":69BE4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":6B2EE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":6C9F81
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":6E101B
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":6F80B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":70F14F
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":7261E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":73D283
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":75431D
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":76B3B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":782451
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":7994EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":7B0585
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":7C761F
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":7DE6B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":7F5753
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":80C7ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":823887
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":83A921
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":8519BB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16g 
      Left            =   12285
      Top             =   8910
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":868A55
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":868FEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":869589
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":869923
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":869CBD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":86A057
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   10935
      Top             =   8910
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
            Picture         =   "MAIN.frx":86A3F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":86AE03
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":86B815
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":86BBAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":86BF49
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":86C2E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":86C67D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":86D08F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":86DAA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":86E4B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":86EEC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":86F8D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":8702E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":870CFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MAIN.frx":871297
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10785
      Left            =   19320
      ScaleHeight     =   10785
      ScaleWidth      =   2445
      TabIndex        =   4
      Top             =   1245
      Width           =   2445
      Begin PRD.ACPRibbon ACPMenu 
         Height          =   2130
         Left            =   135
         TabIndex        =   22
         Top             =   11025
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   3757
      End
      Begin VB.ComboBox CboPlant 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         ItemData        =   "MAIN.frx":871833
         Left            =   135
         List            =   "MAIN.frx":871840
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   450
         Width           =   2220
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   240
         Left            =   135
         TabIndex        =   19
         Top             =   6210
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox txtStdAktual 
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
         Height          =   405
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txtStdParameter 
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
         Height          =   405
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txtSkillMatrix 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   2430
         Width           =   2220
      End
      Begin VB.ComboBox cboMachine 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         ItemData        =   "MAIN.frx":87185B
         Left            =   1260
         List            =   "MAIN.frx":87185D
         TabIndex        =   10
         Text            =   "1"
         Top             =   1215
         Width           =   1095
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
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0"
         Top             =   4320
         Width           =   2220
      End
      Begin VB.TextBox txtIdletime 
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
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0"
         Top             =   5355
         Width           =   2220
      End
      Begin VB.Label lblMachine 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   135
         TabIndex        =   20
         Top             =   1800
         Width           =   2220
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "ACTUAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1260
         TabIndex        =   17
         Top             =   3330
         Width           =   1095
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "ENG-STD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   135
         TabIndex        =   16
         Top             =   3330
         Width           =   1095
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "STANDARD PARAMETER"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   135
         TabIndex        =   15
         Top             =   3015
         Width           =   2220
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "PLANT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   135
         TabIndex        =   13
         Top             =   90
         Width           =   2220
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "SKILL MATRIX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   135
         TabIndex        =   12
         Top             =   2160
         Width           =   2220
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "No. Machine"
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
         Height          =   510
         Left            =   135
         TabIndex        =   9
         Top             =   1215
         Width           =   1005
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "SHOT MACHINE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   135
         TabIndex        =   8
         Top             =   4050
         Width           =   2220
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "IDLE TIME"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   135
         TabIndex        =   6
         Top             =   5040
         Width           =   2220
      End
   End
   Begin b8Controls4.b8ClientWin b8CW 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   31
      Top             =   870
      Width           =   21765
      _ExtentX        =   38391
      _ExtentY        =   661
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu smSetting 
         Caption         =   "Settings"
      End
      Begin VB.Menu smAdjustShot 
         Caption         =   "Adjust Shot"
      End
      Begin VB.Menu smAdjustNG 
         Caption         =   "Adjust NG"
      End
      Begin VB.Menu smFile1 
         Caption         =   "-"
      End
      Begin VB.Menu smChangeUser 
         Caption         =   "Change User"
      End
      Begin VB.Menu smLoginUser2 
         Caption         =   "Login User 2"
      End
      Begin VB.Menu smFile2 
         Caption         =   "-"
      End
      Begin VB.Menu smExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu smTransactions 
      Caption         =   "Transactions"
      Begin VB.Menu smDashboard 
         Caption         =   "Dashboard"
      End
      Begin VB.Menu smWI 
         Caption         =   "WI/CCP/SP"
      End
      Begin VB.Menu smInfo 
         Caption         =   "Info"
      End
      Begin VB.Menu smCall 
         Caption         =   "Call/SMS"
      End
      Begin VB.Menu smTools1 
         Caption         =   "-"
      End
      Begin VB.Menu smProdResult 
         Caption         =   "Prod Result"
      End
      Begin VB.Menu smNGReject 
         Caption         =   "NG/Reject"
      End
      Begin VB.Menu smIdleTime 
         Caption         =   "Idle Time"
      End
      Begin VB.Menu smTools2 
         Caption         =   "-"
      End
      Begin VB.Menu smSkillGeneral 
         Caption         =   "Skill General"
      End
      Begin VB.Menu smSkillProduct 
         Caption         =   "Skill Product"
      End
      Begin VB.Menu smTools3 
         Caption         =   "-"
      End
      Begin VB.Menu smAbsensi 
         Caption         =   "Absensi Karyawan"
      End
      Begin VB.Menu smMonitoringLeader 
         Caption         =   "Monitoring Leaders"
      End
      Begin VB.Menu smTroubleMachine 
         Caption         =   "Trouble Machine"
      End
      Begin VB.Menu smChartMonitoring 
         Caption         =   "Chart Monitoring Leader"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Report"
      Begin VB.Menu smDailyReport 
         Caption         =   "Daily Report"
      End
      Begin VB.Menu smDailyYieldReport 
         Caption         =   "Daily Yield Report "
      End
      Begin VB.Menu smDatabyProduct 
         Caption         =   "Data by Product"
      End
      Begin VB.Menu smYieldbyMaterial 
         Caption         =   "Yield by Material"
      End
      Begin VB.Menu smSetupMold 
         Caption         =   "Setup Mold"
      End
      Begin VB.Menu smSetupMoldUser 
         Caption         =   "Setup Mold Per User"
      End
      Begin VB.Menu smBreakdownNG 
         Caption         =   "Breakdown NG"
      End
      Begin VB.Menu smRangkumanCustomer 
         Caption         =   "Rangkuman Customer"
      End
      Begin VB.Menu smLabelReport 
         Caption         =   "Label Report"
      End
      Begin VB.Menu smreport2 
         Caption         =   "-"
      End
      Begin VB.Menu smSmListWI 
         Caption         =   "List WI/CP/PS"
      End
      Begin VB.Menu smDataIdle 
         Caption         =   "Data Idle"
      End
      Begin VB.Menu smDataNG 
         Caption         =   "Data NG"
      End
      Begin VB.Menu smMonitoringMachine 
         Caption         =   "Monitoring Machine"
      End
      Begin VB.Menu smParameterStandard 
         Caption         =   "Parameter Standard"
      End
      Begin VB.Menu smTotalYield 
         Caption         =   "Total Yield Monitoring"
      End
      Begin VB.Menu smYieldReport 
         Caption         =   "Total Yield Report"
      End
      Begin VB.Menu smLoginUsers 
         Caption         =   "Login Users"
      End
      Begin VB.Menu smPanggilanTeknisi 
         Caption         =   "Panggilan Teknisi"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu smAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Visible         =   0   'False
      Begin VB.Menu mnuRACN 
         Caption         =   "Create New Entry"
      End
      Begin VB.Menu mnuRAES 
         Caption         =   "Edit Selected"
      End
      Begin VB.Menu mnuRADS 
         Caption         =   "Delete Selected"
      End
      Begin VB.Menu mnuRARR 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuRAP 
         Caption         =   "Export"
      End
      Begin VB.Menu mnuRAC 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuSRC 
         Caption         =   "Search"
      End
   End
End
Attribute VB_Name = "MAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim cursor_pos As POINTAPI
Dim resize_down     As Boolean
Dim show_mnu        As Boolean
Dim pos_num         As Integer
Dim Theme           As Integer
Public CloseMe      As Boolean
Dim clsDS2          As New clsDS2


Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim date1 As Date, date2 As Date
Dim seconds As Integer, minutes As Integer, hours As Integer



Dim X As Long
Dim Y As Long, Yi As Long, Yx As Long
Dim xTime As Integer

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim x1 As Boolean

'Control Procedures
'-----------------------------------------------------------
Private Sub b8CW_FormTabClick(ByVal sFormName As String, ByVal Index As Integer)
    ActivateMDIChildForm sFormName
End Sub

Public Sub RemoveChild(ByVal sFormName As String)
    'remove form
    'MAIN.b8CW.Visible = False
    Me.b8CW.RemoveChildWindow sFormName
End Sub

Public Sub ActivateChild(ByRef CFrm As Form)
    MAIN.b8CW.Visible = True
    'activate form
    Me.b8CW.SetActiveWindow CFrm.Name
End Sub

Private Sub cboMachine_Click()
    'ReadINI("SETTING", "MACHINE", App.Path & "\Settings.ini")
    Call WriteINI("SETTING", "MACHINE", cboMachine.text, App.Path & "\Settings.ini")
    
    NoMesin = ReadINI("SETTING", "MACHINE", App.Path & "\Settings.ini")
    Call LoadProduct
    Call TotalShot(p_eng_product_1)
    
    Call TotalIdle
    Call SkillMatrix
    Call StandardParameter
    Call ActualParameter
    
    Unload FrmDashboard
    lblMachine.Caption = p_machine_name
    
End Sub



Private Sub CboPlant_Click()

    Call WriteINI("Server", "PLANT", CboPlant.text, App.Path & "\Database.ini")
    p_plant_mark = ReadINI("Server", "PLANT", App.Path & "\Database.ini")
    
    Set CN = Nothing
    
    If Connected2DB = False Then Unload Me: Exit Sub
    
    Call LoadProduct
    Call TotalShot(p_eng_product_1)
    
    Call TotalIdle
    Call SkillMatrix
    Call StandardParameter
    Call ActualParameter
    
    Unload FrmDashboard
    lblMachine.Caption = p_machine_name
    
End Sub






Private Sub Label16_DblClick()

        Shape1.FillColor = vbGreen
        If p_status_prod_1 = True Then
            Add_counter p_eng_product_1
        End If
        
        If p_status_prod_2 = True Then
            Add_counter p_eng_product_2
        End If

        If p_status_prod_3 = True Then
            Add_counter p_eng_product_3
        End If

        If p_status_prod_4 = True Then
            Add_counter p_eng_product_4
        End If
        
        Call TotalShot(p_eng_product_1)
        lblCT.Caption = DateDiff("s", Format(Label1.Caption, "hh:mm:ss"), Label2.Caption)
        Label1.Caption = Format(Now, "hh:mm:ss")
        
        Shape1.FillColor = vbRed
        

End Sub


Private Sub lbloff_Change()
On Error GoTo ErrHandler

    If lbloff.Caption = PortOn Then
        Shape1.FillColor = vbGreen
        If p_status_prod_1 = True Then
            Add_counter p_eng_product_1
        End If
        If p_status_prod_2 = True Then
            Add_counter p_eng_product_2
        End If

        If p_status_prod_3 = True Then
            Add_counter p_eng_product_3
        End If
        
        If p_status_prod_4 = True Then
            Add_counter p_eng_product_4
        End If
        
        lblCT.Caption = DateDiff("s", Format(Label1.Caption, "hh:mm:ss"), Label2.Caption)
        
        Label1.Caption = Format(Now, "hh:mm:ss")
        
    Else
        Shape1.FillColor = vbRed
    End If
    
    Call TotalShot(p_eng_product_1)
    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation


End Sub



Private Sub MDIForm_Activate()
On Error Resume Next
 X = 0
 Y = 0
 Yi = 0
 
 
 
If END_APP = True Then End: Exit Sub

End Sub

Private Sub MDIForm_Load()
On Error GoTo ErrHandler

'# SET Theme
Dim iStyle As String
iStyle = ReadINI("THEME", "Style", App.Path & "\Setting.ini")
If iStyle = "" Then
    ACPMenu.Theme = 1
Else
    ACPMenu.Theme = 1
End If
MAIN.Picture = ACPMenu.LoadBackground

'# Set ImageList to use for icons
ACPMenu.ImageList = ImageList2 'i32x32
Picture1.BackColor = ACPMenu.BackColor

ACPMenu.Align = vbAlignTop

ACPMenu.ButtonCenter = True

MAIN.BackColor = ACPMenu.BackColor
'# Set Circle Menu Button Picture with Index of image on imagelist
ACPMenu.Icon = 104


If ReadINI("SERVER", "MACHINE", App.Path & "\Database.ini") = "" Then
    Display_Mesin = False
Else
    Display_Mesin = ReadINI("SERVER", "MACHINE", App.Path & "\Database.ini")
End If

If ReadINI("SERVER", "SETTING", App.Path & "\Database.ini") = "" Then
    Display_Setting = False
Else
    Display_Setting = ReadINI("SERVER", "SETTING", App.Path & "\Database.ini")
End If

    If Display_Mesin = False Then
    
        MAIN.Toolbar1.Buttons(2).Enabled = False
        MAIN.Toolbar1.Buttons(3).Enabled = False
        MAIN.Toolbar1.Buttons(4).Enabled = False
        
        MAIN.Toolbar1.Buttons(6).Enabled = False
        MAIN.Toolbar1.Buttons(7).Enabled = False
        MAIN.Toolbar1.Buttons(8).Enabled = False
    
        MAIN.Toolbar1.Buttons(10).Enabled = False
        MAIN.Toolbar1.Buttons(11).Enabled = False

        MAIN.Toolbar1.Buttons(13).Enabled = False
        MAIN.Toolbar1.Buttons(15).Enabled = False
        MAIN.Toolbar1.Buttons(16).Enabled = False
    
    End If

    If Display_Setting = False Then
        MAIN.smAdjustShot.Enabled = False
        MAIN.smAdjustNG.Enabled = False
        MAIN.smSetting.Enabled = False
    End If

    MAIN.Caption = "PMS - Version " & App.Major & "." & App.Minor & " (Build: " & App.Revision & ")"

    NoMesin = ReadINI("SETTING", "MACHINE", App.Path & "\Settings.ini")

    PortOn = ReadINI("SETTING", "SERIAL", App.Path & "\Settings.ini")
    IdleOn = ReadINI("SETTING", "IDLEON", App.Path & "\Settings.ini")
    LinkUpdate = ReadINI("SETTING", "LINKUPDATE", App.Path & "\Settings.ini")
    MaxShot = ReadINI("SETTING", "MAXSHOT", App.Path & "\Settings.ini")

    Dim i As Integer
    For i = 1 To 80
        cboMachine.AddItem i
    Next i
    
    cboMachine.text = ReadINI("SETTING", "MACHINE", App.Path & "\Settings.ini")
    p_plant_mark = ReadINI("Server", "PLANT", App.Path & "\Database.ini")

    If p_plant_mark = "TSSI" Then
        p_sys_plant = "3,4"
    ElseIf p_plant_mark = "Techno" Then
        p_sys_plant = "2"
    ElseIf p_plant_mark = "Cempaka" Then
        p_sys_plant = "6"
    End If
    
    lblOn.Caption = PortOn
    
    PortAddress = &H379
    
    Me.Show

    If Connected2DB = False Then Unload Me: Exit Sub
    'If Connected2SIP = False Then Unload Me: Exit Sub
    
    frmLogin.Show vbModal
    
    Call HideMenu
    
    Call GetShift
    
    Call LoadProduct
    
    CboPlant.text = p_plant_mark
    
    lblMachine.Caption = p_machine_name
    Label1.Caption = Format(Now, "hh:mm:ss")
    
    Timer1.Enabled = True
    Timer2.Enabled = True
    Timer3.Enabled = True
    TimerPort.Enabled = True
    
    Call TotalShot(p_eng_product_1)

    bIdle = False
    bformIdle = False

    If ReadINI("SETTING", "SHOWINFO", App.Path & "\Settings.ini") <> "" Then
        If ReadINI("SETTING", "SHOWINFO", App.Path & "\Settings.ini") = 1 Then
            xTime = ReadINI("SETTING", "INFO", App.Path & "\Settings.ini")
            TmrInformasi.Enabled = True
        Else
            TmrInformasi.Enabled = False
        End If
    Else
        TmrInformasi.Enabled = False
    End If
    
    Call SkillMatrix
    Call StandardParameter
    Call ActualParameter
    
    DisplayBusinessInfo
    Me.b8CW.SetActiveWindow ""
    Me.b8CW.Visible = False
    
    
    
    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation

End Sub

Private Sub HideMenu()

    If ReadINI("SETTING", "SHOWSENSOR", App.Path & "\Settings.ini") = "" Then
        ShowSensor = True
        frameSensor.Visible = True
    Else
        ShowSensor = ReadINI("SETTING", "SHOWSENSOR", App.Path & "\Settings.ini")
        If ShowSensor = True Then
            frameSensor.Visible = True
        Else
            frameSensor.Visible = False
        End If
    End If
    
    If ReadINI("SETTING", "SHOWWI", App.Path & "\Settings.ini") <> "" Then
        ShowWI = ReadINI("SETTING", "SHOWWI", App.Path & "\Settings.ini")
        If ShowWI = True Then
            LoadForm frmProduction
        End If
    End If

    If ReadINI("SETTING", "STOOLBAR", App.Path & "\Settings.ini") <> "" Then
        ShowToolbar = ReadINI("SETTING", "STOOLBAR", App.Path & "\Settings.ini")
    End If

    If ReadINI("SETTING", "ENMACHINE", App.Path & "\Settings.ini") <> "" Then
         eMachine = ReadINI("SETTING", "ENMACHINE", App.Path & "\Settings.ini")
        If eMachine = False Then
            cboMachine.Enabled = False
        Else
            cboMachine.Enabled = True
        End If
    End If

End Sub


Private Sub MDIForm_Initialize()
    On Error Resume Next
        ' this will fail if Comctl not available
        '  - unlikely now though!
        Dim iccex As tagInitCommonControlsEx
        With iccex
            .lngSize = LenB(iccex)
            .lngICC = ICC_USEREX_CLASSES
        End With
        InitCommonControlsEx iccex
End Sub


Public Sub UnloadChilds()
''Unload all active forms
Dim Form As Form
   For Each Form In Forms
      ''Unload all active childs
      If Form.Name <> Me.Name And Form.Name <> "frmProduction" Then Unload Form
   Next Form
   
Set Form = Nothing
End Sub

Private Sub MDIForm_Resize()
    ACPMenu.Resize
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    Call UnloadChilds
    Set MAIN = Nothing
    Set CN = Nothing
    End
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Reply As Integer

Reply = MsgBox("This will terminate the application.Do you want to proceed?", vbExclamation + vbYesNo)

If Reply = vbNo Then
    Cancel = 1
End If

End Sub


Private Sub mnuRACN_Click()
    On Error Resume Next
    ActiveForm.CommandPass "New"
End Sub

Private Sub mnuRADS_Click()
    On Error Resume Next
    ActiveForm.CommandPass "Delete"
End Sub

Private Sub mnuRAES_Click()
    On Error Resume Next
    ActiveForm.CommandPass "Update"
End Sub

Private Sub mnuRAP_Click()
    On Error Resume Next
    ActiveForm.CommandPass "Export"
End Sub

Private Sub mnuRARR_Click()
    On Error Resume Next
    ActiveForm.CommandPass "Refresh"
End Sub

Private Sub mnuRAC_Click()
    On Error Resume Next
    ActiveForm.CommandPass "Close"
End Sub

Private Sub mnuSRC_Click()
    On Error Resume Next
    ActiveForm.CommandPass "Search"
End Sub




Private Sub smAbsensi_Click()
    Call UnloadChilds
    LoadForm frmAbsensi
    
End Sub

Private Sub smAdjustNG_Click()
    Call UnloadChilds
    LoadForm frmAdjustNG
End Sub

Private Sub smAdjustShot_Click()
    Call UnloadChilds
    LoadForm frmAdjustshot
    
End Sub

Private Sub smBreakdownNG_Click()
    Call UnloadChilds
    LoadForm frmBdownNG
End Sub

Private Sub smCall_Click()
    UnloadChilds
    frmCallSMS.Show 1
    
End Sub

Private Sub smChangeUser_Click()
On Error Resume Next
    If MsgBox("Silahkan Logout sebelum mengganti dengan User yang lain, apakah mau diganti ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Dim iSQL As String
    iSQL = "INSERT INTO hrd_login_logs (plant_mark,loc_code,hrd_employee_id,emp_code"
    iSQL = iSQL & " ,tr_date,tr_time,acc_code,created_at,created_by)"
    iSQL = iSQL & " VALUES ('" & p_plant_mark & "','" & NoMesin & "','" & ACTIVE_USER.KODEUSER & "','" & ACTIVE_USER.KODEPIN & "'"
    iSQL = iSQL & " ,'" & Format(Now, "yyyy-mm-dd") & "','" & Format(Now, "hh:mm:ss") & "','2'"
    iSQL = iSQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER.SYSID & "')"
    
    sSQL_Insert iSQL

    Call UnloadChilds
    frmLogin.Show vbModal
End Sub

Private Sub smChartMonitoring_Click()
    LoadForm frmChartMonitoring
End Sub

Private Sub smDailyReport_Click()
    LoadForm frmRptHarian
End Sub

Private Sub smDailyYieldReport_Click()
    LoadForm frmRptYieldHarian
End Sub

Private Sub smDashboard_Click()
    LoadForm FrmDashboard
End Sub

Private Sub smDatabyProduct_Click()
    LoadForm frmDatabyProduct
    
End Sub

Private Sub smDataIdle_Click()
    LoadForm frmDataIdle
    
End Sub

Private Sub smDataNG_Click()
    LoadForm frmDataNG
    
End Sub

Private Sub smExit_Click()
    Unload Me
End Sub

Private Sub smIdleTime_Click()
    frmIdleTime.Show 1
End Sub

Private Sub smInfo_Click()
    frmInformasi.Show 1
    
End Sub

Private Sub smLabelReport_Click()
    LoadForm frmRptLabel
    
End Sub

Private Sub smLoginUser2_Click()
    frmLogin2.Show 1
    
End Sub

Private Sub smLoginUsers_Click()
    LoadForm frmLoguser
End Sub

Private Sub smMonitoringLeader_Click()
    LoadForm frmMonitoringLeader
End Sub

Private Sub smMonitoringMachine_Click()
    frmInfo.Show 1
    
End Sub

Private Sub smNGReject_Click()
    frmNg.Show vbModeless
End Sub

Private Sub smPanggilanTeknisi_Click()
    LoadForm frmRptCall
End Sub

Private Sub smParameterStandard_Click()
    LoadForm frmParameter
    
End Sub

Private Sub smProdResult_Click()
    frmInputResult.Show vbModeless
End Sub

Private Sub smRangkumanCustomer_Click()
    LoadForm frmRangkumanCust
    
End Sub

Private Sub smSetting_Click()
    frmSettings.Show 1
   
End Sub

Private Sub smSetupMold_Click()
    LoadForm frmSetupMold
    
End Sub

Private Sub smSetupMoldUser_Click()
    LoadForm frmSetupMoldUser
End Sub

Private Sub smSkillGeneral_Click()
    RptSkillGeneral.Show vbModeless
    
End Sub

Private Sub smSkillProduct_Click()
    RptSkillMatrik.Show vbModeless
End Sub

Private Sub smSmListWI_Click()
    LoadForm frmDataPDF
    
End Sub

Private Sub smTotalYield_Click()
    frmTotalYield.Show

End Sub

Private Sub smTroubleMachine_Click()
    LoadForm frmTroubleMachine
End Sub

Private Sub smWI_Click()
    LoadForm frmProduction
End Sub

Private Sub smYieldbyMaterial_Click()
    LoadForm frmProdYieldByMtrl
    
End Sub

Private Sub smYieldReport_Click()
    frmYieldReport.Show 1
End Sub

Private Sub Timer1_Timer()
    'Check Label
    Label2.Caption = Format(Now, "hh:mm:ss")
    Label4.Caption = DateDiff("s", Format(Label1.Caption, "hh:mm:ss"), Label2.Caption)
    
    If bformIdle = False Then
        If bIdle = False Then
            If IdleOn = True Then
                If Val(Label4.Caption) = (Val(Round(p_cycle_time_1, 0)) * 5) Then
                    bIdle = True
                    frmIdleTime.Show 1
                End If
            End If
        End If
    End If
    
End Sub



Private Sub Timer2_Timer()
    Yi = Yi + 1
    If Yi >= 10 Then
        Call SkillMatrix
        Call StandardParameter
        Call ActualParameter

        Yi = 0
    End If
    
End Sub

Private Sub Timer3_Timer()
    ProgressBar1.Max = 120
    Yx = Yx + 1
    If Yx >= 120 Then
        If p_status_prod_1 = True Then
            AddCounterAwal p_eng_product_1
        End If
        If p_status_prod_2 = True Then
            AddCounterAwal p_eng_product_2
        End If
        
        'Call TotalShot(p_eng_product_1)

        Call TotalIdle
    
        Call GetShift
        Call LoadProduct
        

        Yx = 0
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Value = ProgressBar1.Value + 1
    End If
End Sub


    
Private Sub TimerPort_Timer()
    'ShowSensor = ReadINI("SETTING", "SHOWSENSOR", App.Path & "\Settings.ini")
    If ShowSensor = True Then
        lbloff.Caption = PortIn(PortAddress)
    End If
End Sub


Private Sub Add_counter(eng_prod As Variant)
On Error GoTo ErrHandler

    Dim Rs As New Recordset
    Dim sSQL, iSQL As String
    Dim Counter As Variant
    Dim eng_prod_1 As String, eng_prod_2 As String

    Counter = 0
    Rs.CursorLocation = adUseClient
    
    sSQL = "select a.counter_ok from sip_production.prod_runnings a where "
    sSQL = sSQL & " a.plant_mark = '" & p_plant_mark & "' "
    sSQL = sSQL & " and a.prod_machine_id = '" & p_prod_machine_id & "'"
    sSQL = sSQL & " and a.mkt_customer_id = '" & p_mkt_customer_id & "'"
    sSQL = sSQL & " and a.eng_product_id = '" & eng_prod & "'"
    sSQL = sSQL & " and a.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "'"
    sSQL = sSQL & " and a.period_hour = '" & Format(Now, "HH") & "'"
    
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open sSQL, CN, adOpenStatic, adLockReadOnly
    
    If Rs.RecordCount < 1 Then
        iSQL = "insert into sip_production.prod_runnings (plant_mark,prod_machine_id,mkt_customer_id,eng_product_id,"
        iSQL = iSQL & " date,period_shift,period_hour,counter_ok,created_at,created_by,operator_1,operator_2) values"
        iSQL = iSQL & " ('" & p_plant_mark & "','" & p_prod_machine_id & "','" & p_mkt_customer_id & "','" & eng_prod & "'"
        iSQL = iSQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & Format(p_shift, "yyyy-mm-dd") & "','" & Format(Now, "HH") & "','" & 1 & "'"
        iSQL = iSQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER.KODEUSER & "','" & ACTIVE_USER.KODEUSER & "','" & ACTIVE_USER_2.KODEUSER & "')"
        
        'sSQL_Insert iSQL
        
        CN.BeginTrans
        CN.Execute iSQL
        CN.CommitTrans

    Else
        Counter = Val(Rs.Fields("counter_ok")) + 1
        iSQL = "update sip_production.prod_runnings set counter_ok = '" & Counter & "',operator_1 = '" & ACTIVE_USER.KODEUSER & "',operator_2 = '" & ACTIVE_USER_2.KODEUSER & "',"
            iSQL = iSQL & " date = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' where "
            iSQL = iSQL & " plant_mark = '" & p_plant_mark & "' "
            iSQL = iSQL & " and prod_machine_id = '" & p_prod_machine_id & "'"
            iSQL = iSQL & " and mkt_customer_id = '" & p_mkt_customer_id & "'"
            iSQL = iSQL & " and eng_product_id = '" & eng_prod & "'"
            iSQL = iSQL & " and period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "'"
            iSQL = iSQL & " and period_hour = '" & Format(Now, "HH") & "'"
    
        'sSQL_Update sSQL
        
        CN.BeginTrans
        CN.Execute iSQL
        CN.CommitTrans

    End If
    
    Set Rs = Nothing
    
Exit Sub
ErrHandler:
    If Err.Number <> 0 Then
        CN.RollbackTrans
    End If
End Sub




Private Sub TmrInformasi_Timer()
    Y = Y + 1
    
    If Y >= xTime Then
        If formNG = False Then
            If formIdle = False Then
                frmInformasi.Show 1
                Y = 0
            End If
        End If
    End If

End Sub



Private Sub TotalShot(sProd As String)
    Dim Rs As New ADODB.Recordset
    Dim sSQL As String

    Rs.CursorLocation = adUseClient

    sSQL = "SELECT a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift,SUM(a.counter_ok) AS total_shot"
    sSQL = sSQL & " FROM sip_production.prod_runnings a"
    sSQL = sSQL & " where a.plant_mark = '" & p_plant_mark & "' and"
    sSQL = sSQL & " a.prod_machine_id = '" & p_prod_machine_id & "' and"
    sSQL = sSQL & " a.mkt_customer_id = '" & p_mkt_customer_id & "' and"
    sSQL = sSQL & " a.eng_product_id = '" & sProd & "' and"
    sSQL = sSQL & " a.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "'"
    sSQL = sSQL & " GROUP BY a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift"

    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open sSQL, CN, adOpenStatic, adLockReadOnly
    
    If Rs.RecordCount > 0 Then
        txtShot_1 = IIf(IsNull(Rs.Fields("total_shot")), "0", Rs.Fields("total_shot"))
        
    End If
    
    Label3.Caption = p_cycle_time_1
    
    Set Rs = Nothing
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
        Case "dashboard"
            LoadForm FrmDashboard
        Case "WI/CCP/SP"
            LoadForm frmProduction
        Case "Info"
            frmInformasi.Show vbModeless
        Case "Call/SMS"
            frmCallSMS.Show vbModeless
        Case "Prod Result"
            frmInputResult.Show vbModeless
        Case "NG/Reject"
            frmNg.Show vbModeless
        Case "Idle Time"
            frmIdleTime.Show 1
        Case "Monitoring Leader"
            LoadForm frmMonitoringLeader
        Case "Trouble Machine"
            LoadForm frmTroubleMachine

      
        Case "Change User"
            If MsgBox("Silahkan Logout sebelum mengganti dengan User yang lain, apakah mau diganti ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            
            Dim iSQL As String
            iSQL = "INSERT INTO hrd_login_logs (plant_mark,loc_code,hrd_employee_id,emp_code"
            iSQL = iSQL & " ,tr_date,tr_time,acc_code,created_at,created_by)"
            iSQL = iSQL & " VALUES ('" & p_plant_mark & "','" & NoMesin & "','" & ACTIVE_USER.KODEUSER & "','" & ACTIVE_USER.KODEPIN & "'"
            iSQL = iSQL & " ,'" & Format(Now, "yyyy-mm-dd") & "','" & Format(Now, "hh:mm:ss") & "','2'"
            iSQL = iSQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER.SYSID & "')"
            
            sSQL_Insert iSQL
            frmLogin.Show vbModal
            
        Case "Login User 2"
            frmLogin2.Show 1
            
    End Select
End Sub





Private Sub TotalIdle()
On Error GoTo ErrHandler

    Dim Rs As New Recordset
    Dim sSQL As String
  
    Rs.CursorLocation = adUseClient
    
    sSQL = "select a.plant_mark,a.prod_machine_id,a.period_shift,SEC_TO_TIME(sum(time_to_sec(idle_time))) as losstime"
    sSQL = sSQL & " from sip_production.prod_machine_idles a"
    sSQL = sSQL & " where a.plant_mark = '" & p_plant_mark & "' and"
    sSQL = sSQL & " a.prod_machine_id = '" & p_prod_machine_id & "' and"
    sSQL = sSQL & " a.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "'"
    sSQL = sSQL & " group by a.plant_mark,a.prod_machine_id,a.period_shift"
    
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open sSQL, CN, adOpenStatic, adLockReadOnly
    
    If Rs.RecordCount > 0 Then
        txtIdletime.text = IIf(IsNull(Rs.Fields("losstime")), "0", Rs.Fields("losstime"))
        txtIdletime.text = Format(txtIdletime.text, "hh:mm:ss")
    End If
    Set Rs = Nothing

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
    
End Sub

Private Sub SkillMatrix()
On Error GoTo ErrHandler

    Dim Rs As New Recordset
    Dim sSQL As String
  
    Rs.CursorLocation = adUseClient
    
    sSQL = "select a.id,b.nik,b.nama_karyawan,c.prod_skill_product_id,c.eng_product_id,c.internal_part_id,c.product_name"
    sSQL = sSQL & " from prod_skill_products a"
    sSQL = sSQL & " INNER JOIN (select emp.id, emp.nik,emp.name as nama_karyawan,dep.name as departement,"
        sSQL = sSQL & " pos.name as positions from hrd_employees emp"
        sSQL = sSQL & " inner join sys_departments dep on emp.sys_department_id = dep.id"
        sSQL = sSQL & " inner join hrd_positions pos on emp.hrd_position_id = pos.id) b on a.hrd_employee_id = b.id"
    sSQL = sSQL & " INNER JOIN (SELECT skill.id,skill.prod_skill_product_id,skill.eng_product_id,prd.internal_part_id,"
        sSQL = sSQL & " prd.name as product_name,skill.pg,skill.cv,skill.rw,skill.pl,skill.ng,skill.result,skill.`status`,"
        sSQL = sSQL & " skill.created_at,skill.created_by,crt_sys.name as created_name,"
        sSQL = sSQL & " skill.approve_1_at,skill.approve_1_by,app1_sys.name as app1_name"
        sSQL = sSQL & " FROM prod_skill_product_items skill"
        sSQL = sSQL & " LEFT JOIN eng_products prd on skill.eng_product_id = prd.id"
        sSQL = sSQL & " LEFT JOIN sys_accounts crt_sys on skill.created_by = crt_sys.id"
        sSQL = sSQL & " LEFT JOIN sys_accounts app1_sys on skill.approve_1_by = app1_sys.id) c"
    sSQL = sSQL & " on a.id = c.prod_skill_product_id where b.nik = '" & Mid(ACTIVE_USER.USERNAME, 2, 7) & "'"
    sSQL = sSQL & " and c.internal_part_id = '" & p_int_part_1 & "'"
    
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open sSQL, CN, adOpenStatic, adLockReadOnly
    
    If Rs.RecordCount > 0 Then
        txtSkillMatrix.text = "ADA"
        txtSkillMatrix.BackColor = vbGreen
    Else
        txtSkillMatrix.text = "BELUM ADA"
        txtSkillMatrix.BackColor = vbRed
    End If
    Set Rs = Nothing

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
    
End Sub


Private Sub AddCounterAwal(eng_prod As Variant)
On Error GoTo ErrHandler

    Dim Rs As New Recordset
    Dim sSQL, iSQL As String
    Dim eng_prod_1 As String, eng_prod_2 As String
    
    

    Rs.CursorLocation = adUseClient
    
    sSQL = "select a.* from sip_production.prod_runnings a where "
    sSQL = sSQL & " a.plant_mark = '" & p_plant_mark & "' "
    sSQL = sSQL & " and a.prod_machine_id = '" & p_prod_machine_id & "'"
    sSQL = sSQL & " and a.mkt_customer_id = '" & p_mkt_customer_id & "'"
    sSQL = sSQL & " and a.eng_product_id = '" & eng_prod & "'"
    sSQL = sSQL & " and a.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "'"
    sSQL = sSQL & " and a.period_hour = '" & Format(Now, "HH") & "'"
    
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open sSQL, CN, adOpenStatic, adLockReadOnly
    
    If Rs.RecordCount < 1 Then
        iSQL = "insert into sip_production.prod_runnings (plant_mark,prod_machine_id,mkt_customer_id,eng_product_id,"
        iSQL = iSQL & " date,period_shift,period_hour,counter_ok,created_at,created_by,operator_1,operator_2) values"
        iSQL = iSQL & " ('" & p_plant_mark & "','" & p_prod_machine_id & "','" & p_mkt_customer_id & "','" & eng_prod & "'"
        iSQL = iSQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & Format(p_shift, "yyyy-mm-dd") & "','" & Format(Now, "HH") & "','" & 0 & "'"
        iSQL = iSQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER.KODEUSER & "','" & ACTIVE_USER.KODEUSER & "','" & ACTIVE_USER_2.KODEUSER & "')"
        
        'sSQL_Insert iSQL
        CN.BeginTrans
        CN.Execute iSQL
        CN.CommitTrans

    End If
    
    Set Rs = Nothing
    
Exit Sub
ErrHandler:
    If Err.Number <> 0 Then
        CN.RollbackTrans
    End If
End Sub


Private Sub StandardParameter()
On Error GoTo ErrHandler

    Dim Rs As New Recordset
    Dim sSQL As String
  
    Rs.CursorLocation = adUseClient
    
    If cboMachine.text = "35" Or cboMachine.text = "36" Then
    
    sSQL = "SELECT a.prod_machine_id,a.eng_product_id"
    sSQL = sSQL & " FROM eng_blow_paramset_standards a"
    sSQL = sSQL & " Where a.prod_machine_id = '" & p_prod_machine_id & "'"
    sSQL = sSQL & " and a.eng_product_id = '" & p_eng_product_1 & "'"
    sSQL = sSQL & " ORDER BY rev DESC LIMIT 1"
    
    Else
    
    sSQL = "SELECT a.prod_machine_id,a.eng_product_id"
    sSQL = sSQL & " FROM eng_paramset_standards a"
    sSQL = sSQL & " Where a.prod_machine_id = '" & p_prod_machine_id & "'"
    sSQL = sSQL & " and a.eng_product_id = '" & p_eng_product_1 & "'"
    sSQL = sSQL & " ORDER BY rev DESC LIMIT 1"
    
    End If
    
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open sSQL, CN, adOpenStatic, adLockReadOnly
    
    If Rs.RecordCount > 0 Then
        txtStdParameter.text = "ADA"
        txtStdParameter.BackColor = vbGreen
    Else
        txtStdParameter.text = "BLM ADA"
        txtStdParameter.BackColor = vbRed
    End If
    Set Rs = Nothing

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
    
End Sub

Private Sub ActualParameter()
On Error GoTo ErrHandler

    Dim Rs As New Recordset
    Dim sSQL As String
  
    Rs.CursorLocation = adUseClient

    sSQL = "SELECT a.sys_plant_id, a.mkt_customer_id,"
    sSQL = sSQL & " a.prod_machine_id,a.prod_machine_id,a.eng_product_id,a.date"
    sSQL = sSQL & " FROM eng_paramset_actuals a"
    sSQL = sSQL & " Where a.prod_machine_id = '" & p_prod_machine_id & "'"
    sSQL = sSQL & " and a.eng_product_id = '" & p_eng_product_1 & "'"
    sSQL = sSQL & " and a.date = '" & Format(p_shift, "yyyy-mm-dd") & "'"
    
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open sSQL, CN, adOpenStatic, adLockReadOnly
    
    If Rs.RecordCount > 0 Then
        txtStdAktual.text = "SESUAI"
        txtStdAktual.BackColor = vbGreen
    Else
        txtStdAktual.text = "BLM INPUT"
        txtStdAktual.BackColor = vbRed
    End If
    Set Rs = Nothing

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
    
End Sub



Private Sub txtStdAktual_DblClick()
    RptActParameter.Show 1
End Sub

Private Sub txtStdParameter_DblClick()
On Error Resume Next
Dim qSQL As String

qSQL = "SELECT a.*,b.name AS plant_name,c.name AS customer_name,d.number AS machine_number, "
qSQL = qSQL & " d.name AS machine_name,d.tonnage,e.internal_part_id,e.name AS product_name,"
qSQL = qSQL & " e.customer_part_number, e.customer_part_name,e.cavity,e.eng_material_id,"
qSQL = qSQL & " e.eng_color_id,e.weight_gr,e.weight_shot_gr, e.weight_runner_gr ,"
qSQL = qSQL & " e.prod_yield,e.material_name,e.color_name FROM eng_paramset_standards a"
qSQL = qSQL & " LEFT JOIN sys_plants b ON a.sys_plant_id = b.id"
qSQL = qSQL & " LEFT JOIN mkt_customers c ON a.mkt_customer_id = c.id"
qSQL = qSQL & " LEFT JOIN prod_machines d ON a.sys_plant_id = d.id"
qSQL = qSQL & " LEFT JOIN (SELECT a.*,b.name AS material_name, c.name AS color_name FROM eng_products a"
                qSQL = qSQL & " LEFT JOIN eng_materials b ON a.eng_material_for_label_id = b.id"
                qSQL = qSQL & " LEFT JOIN eng_colors c ON a.eng_color_id = c.id) e ON a.eng_product_id = e.id "
qSQL = qSQL & " WHERE a.mkt_customer_id = '" & p_mkt_customer_id & "' AND"
qSQL = qSQL & " a.prod_machine_id = '" & p_prod_machine_id & "' AND"
qSQL = qSQL & " a.eng_product_id = '" & p_eng_product_1 & "'"
qSQL = qSQL & " ORDER BY a.rev DESC LIMIT 1"


Set RS_PRINT = New ADODB.Recordset
If RS_PRINT.State = adStateOpen Then RS_PRINT.Close
RS_PRINT.Open qSQL, CN, adOpenStatic, adLockReadOnly
With RptStdParameter
    .DTRpt.Recordset = RS_PRINT
    .lblDate.Caption = Now
    .lblCompany.Caption = ACTIVE_COMPANY.Perusahaan
    .lblAlamat.Caption = ACTIVE_COMPANY.Alamat
    .Show
End With
End Sub


Private Sub DisplayBusinessInfo()
On Error Resume Next
If p_plant_mark = "TSSI" Then
    With ACTIVE_COMPANY
        .IDPerusahaan = "3"
        .Perusahaan = "PT Tri-Saudara Sentosa Industri"
        .Alamat = "Jl. Pinang Block F17 No.3 Delta Sillicon III Lippo Cikarang-Bekasi, Jawa Barat – Indonesia"
        .NoTelepon = "021-8911 1080"
    End With
ElseIf p_plant_mark = "Techno" Then
    With ACTIVE_COMPANY
        .IDPerusahaan = "2"
        .Perusahaan = "PT Techno Indonesia"
        .Alamat = "Jl. Jati 6 Blok J5 No. 19 Newton Techno Park Lippo Cikarang, Bekasi 17550"
        .NoTelepon = "021-8990 2202"
    End With

ElseIf p_plant_mark = "Cempaka" Then
    With ACTIVE_COMPANY
        .IDPerusahaan = "6"
        .Perusahaan = "PT Techno Indonesia"
        .Alamat = "Jl. Cempaka Block F16 No.30A Delta Sillicon III Lippo Cikarang-Bekasi, Jawa Barat – Indonesia"
        .NoTelepon = "021-8990 2202"
    End With
    
End If

End Sub

