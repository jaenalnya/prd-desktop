VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LvButton.ocx"
Begin VB.Form frmInputResult 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INPUT HASIL PRODUKSI"
   ClientHeight    =   8370
   ClientLeft      =   3735
   ClientTop       =   2280
   ClientWidth     =   17265
   Icon            =   "frmInputResult.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   17265
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboListProduct 
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
      Left            =   10215
      Style           =   2  'Dropdown List
      TabIndex        =   79
      Top             =   135
      Width           =   6855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   855
      ScaleHeight     =   2685
      ScaleWidth      =   5295
      TabIndex        =   58
      Top             =   4455
      Visible         =   0   'False
      Width           =   5325
      Begin lvButton.lvButtons_H cmdNumber 
         Height          =   735
         Index           =   0
         Left            =   135
         TabIndex        =   59
         Top             =   90
         Width           =   870
         _ExtentX        =   1535
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
         Left            =   1170
         TabIndex        =   60
         Top             =   90
         Width           =   870
         _ExtentX        =   1535
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
         Left            =   2205
         TabIndex        =   61
         Top             =   90
         Width           =   870
         _ExtentX        =   1535
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
         Left            =   135
         TabIndex        =   62
         Top             =   945
         Width           =   870
         _ExtentX        =   1535
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
         Left            =   1170
         TabIndex        =   63
         Top             =   945
         Width           =   870
         _ExtentX        =   1535
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
         Left            =   2205
         TabIndex        =   64
         Top             =   945
         Width           =   870
         _ExtentX        =   1535
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
         Left            =   135
         TabIndex        =   65
         Top             =   1800
         Width           =   870
         _ExtentX        =   1535
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
         Left            =   1170
         TabIndex        =   66
         Top             =   1800
         Width           =   870
         _ExtentX        =   1535
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
         Left            =   2205
         TabIndex        =   67
         Top             =   1800
         Width           =   870
         _ExtentX        =   1535
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
         Left            =   3240
         TabIndex        =   68
         Top             =   1800
         Width           =   870
         _ExtentX        =   1535
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
         Height          =   1590
         Index           =   10
         Left            =   4275
         TabIndex        =   69
         Top             =   945
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   2805
         Caption         =   "Enter"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         Left            =   4275
         TabIndex        =   70
         Top             =   90
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   1296
         Caption         =   "Clear"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         Left            =   3240
         TabIndex        =   71
         Top             =   90
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   1296
         Caption         =   "-"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
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
   Begin VB.PictureBox picBarcode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5010
      Left            =   90
      ScaleHeight     =   4980
      ScaleWidth      =   8670
      TabIndex        =   30
      Top             =   2655
      Width           =   8700
      Begin lvButton.lvButtons_H CmdKeyBoard 
         Height          =   375
         Index           =   0
         Left            =   1935
         TabIndex        =   74
         Top             =   855
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   661
         Caption         =   "Keyboard"
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
      Begin VB.TextBox txtMBarcode_2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   6435
         TabIndex        =   3
         Top             =   1260
         Width           =   1770
      End
      Begin VB.TextBox txtMBarcode_1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1935
         TabIndex        =   2
         Top             =   1260
         Width           =   4155
      End
      Begin lvButton.lvButtons_H cmdOk 
         Height          =   555
         Left            =   6750
         TabIndex        =   1
         Top             =   135
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   979
         Caption         =   "SIMPAN"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   33023
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmInputResult.frx":617A
         Enabled         =   0   'False
         cBack           =   16777088
      End
      Begin PRD.Liner Liner3 
         Height          =   30
         Left            =   90
         TabIndex        =   55
         Top             =   1935
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   53
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   2745
         Width           =   2715
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   3195
         Width           =   2715
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   2295
         Width           =   1050
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   2745
         Width           =   1050
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   3195
         Width           =   1050
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   3645
         Width           =   1050
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   7740
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   2295
         Width           =   780
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   7740
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   2745
         Width           =   780
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   7740
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   3195
         Width           =   780
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   7740
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   3645
         Width           =   780
      End
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   2295
         Width           =   2715
      End
      Begin VB.TextBox txtBarcode 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1935
         TabIndex        =   0
         Top             =   135
         Width           =   4560
      End
      Begin lvButton.lvButtons_H CmdKeyBoard 
         Height          =   375
         Index           =   1
         Left            =   6435
         TabIndex        =   75
         Top             =   855
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   661
         Caption         =   "Keyboard"
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
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         TabIndex        =   73
         Top             =   1260
         Width           =   285
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "INPUT MANUAL ID LABEL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   270
         TabIndex        =   72
         Top             =   1260
         Width           =   1635
      End
      Begin VB.Label lblBarcode 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   285
         Left            =   2745
         TabIndex        =   57
         Top             =   675
         Width           =   3210
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   285
         Left            =   3510
         TabIndex        =   56
         Top             =   450
         Width           =   2355
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Part Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         TabIndex        =   54
         Top             =   2790
         Width           =   1815
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Part"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         TabIndex        =   53
         Top             =   3240
         Width           =   1995
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Unik"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4500
         TabIndex        =   52
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4500
         TabIndex        =   51
         Top             =   2790
         Width           =   1095
      End
      Begin VB.Label Label5 
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
         Height          =   330
         Left            =   4500
         TabIndex        =   50
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Operator"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4500
         TabIndex        =   49
         Top             =   3690
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Label"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6615
         TabIndex        =   48
         Top             =   2340
         Width           =   1185
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty / Box"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6615
         TabIndex        =   47
         Top             =   2790
         Width           =   1185
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Cavity"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6615
         TabIndex        =   46
         Top             =   3240
         Width           =   1185
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "No Mesin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6615
         TabIndex        =   45
         Top             =   3690
         Width           =   1185
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Tujuan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         TabIndex        =   44
         Top             =   2340
         Width           =   1815
      End
      Begin VB.Label lbleng_product_id 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   285
         Left            =   2250
         TabIndex        =   43
         Top             =   2790
         Width           =   2355
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "SCAN LABEL BARCODE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   270
         TabIndex        =   31
         Top             =   135
         Width           =   1320
      End
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   510
      Left            =   8280
      TabIndex        =   4
      Top             =   7785
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   900
      Caption         =   "E&XIT"
      CapAlign        =   2
      BackStyle       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16777215
      cFHover         =   16777215
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   4210752
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Peiode Shift"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   90
      TabIndex        =   26
      Top             =   90
      Width           =   4470
      Begin MSComCtl2.DTPicker DTDate 
         Height          =   420
         Left            =   225
         TabIndex        =   76
         Top             =   360
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   741
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   89784323
         CurrentDate     =   43996
      End
      Begin VB.OptionButton optshift1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SHIFT 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         Left            =   225
         TabIndex        =   29
         Top             =   900
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optshift1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SHIFT 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   1
         Left            =   225
         TabIndex        =   28
         Top             =   1350
         Width           =   1815
      End
      Begin VB.OptionButton optshift1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SHIFT 3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   2
         Left            =   225
         TabIndex        =   27
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "* Jika shift 3 pastikan tanggalnya hari ini atau kemarin"
         ForeColor       =   &H000000FF&
         Height          =   465
         Left            =   225
         TabIndex        =   77
         Top             =   2160
         Width           =   3975
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Status Barang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   4635
      TabIndex        =   22
      Top             =   90
      Width           =   4155
      Begin VB.OptionButton optbarang1 
         BackColor       =   &H00FFFFFF&
         Caption         =   " HOLD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   2
         Left            =   360
         TabIndex        =   25
         Top             =   1710
         Width           =   1320
      End
      Begin VB.OptionButton optbarang1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         Left            =   360
         TabIndex        =   24
         Top             =   675
         Width           =   1320
      End
      Begin VB.OptionButton optbarang1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "SISA OK / SISA REWORK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   1
         Left            =   360
         TabIndex        =   23
         Top             =   1215
         Width           =   2805
      End
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   8775
      Top             =   9000
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
            Picture         =   "frmInputResult.frx":C304
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputResult.frx":CD16
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputResult.frx":D728
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputResult.frx":DAC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputResult.frx":DE5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputResult.frx":E1F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputResult.frx":E590
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputResult.frx":EFA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputResult.frx":F9B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputResult.frx":103C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputResult.frx":10DD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputResult.frx":117EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputResult.frx":121FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputResult.frx":12C0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputResult.frx":131AA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4650
      Left            =   90
      TabIndex        =   6
      Top             =   2565
      Width           =   8700
      Begin VB.ComboBox cboProduct 
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
         Left            =   1485
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   225
         Width           =   6855
      End
      Begin VB.TextBox txtInput 
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
         Height          =   495
         Index           =   0
         Left            =   1485
         TabIndex        =   7
         Top             =   855
         Width           =   3975
      End
      Begin lvButton.lvButtons_H CmdKey1 
         Height          =   555
         Index           =   1
         Left            =   1485
         TabIndex        =   8
         Top             =   1395
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         Caption         =   "1"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
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
         cBack           =   16777152
      End
      Begin lvButton.lvButtons_H CmdKey1 
         Height          =   555
         Index           =   2
         Left            =   2295
         TabIndex        =   9
         Top             =   1395
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         Caption         =   "2"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
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
         ImgAlign        =   4
         cBack           =   16777152
      End
      Begin lvButton.lvButtons_H CmdKey1 
         Height          =   555
         Index           =   3
         Left            =   3105
         TabIndex        =   10
         Top             =   1395
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         Caption         =   "3"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
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
         cBack           =   16777152
      End
      Begin lvButton.lvButtons_H CmdKey1 
         Height          =   555
         Index           =   11
         Left            =   1485
         TabIndex        =   11
         Top             =   2655
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         Caption         =   "CLEAR"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         cBack           =   16777152
      End
      Begin lvButton.lvButtons_H CmdKey1 
         Height          =   555
         Index           =   4
         Left            =   3915
         TabIndex        =   12
         Top             =   1395
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         Caption         =   "4"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
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
         cBack           =   16777152
      End
      Begin lvButton.lvButtons_H CmdKey1 
         Height          =   555
         Index           =   5
         Left            =   4725
         TabIndex        =   13
         Top             =   1395
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         Caption         =   "5"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
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
         cBack           =   16777152
      End
      Begin lvButton.lvButtons_H CmdKey1 
         Height          =   555
         Index           =   6
         Left            =   1485
         TabIndex        =   14
         Top             =   2025
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         Caption         =   "6"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
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
         cBack           =   16777152
      End
      Begin lvButton.lvButtons_H CmdKey1 
         Height          =   555
         Index           =   10
         Left            =   4725
         TabIndex        =   5
         Top             =   2655
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         Caption         =   "ENTER"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         cBack           =   16777152
      End
      Begin lvButton.lvButtons_H CmdKey1 
         Height          =   555
         Index           =   7
         Left            =   2295
         TabIndex        =   15
         Top             =   2025
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         Caption         =   "7"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
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
         cBack           =   16777152
      End
      Begin lvButton.lvButtons_H CmdKey1 
         Height          =   555
         Index           =   8
         Left            =   3105
         TabIndex        =   16
         Top             =   2025
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         Caption         =   "8"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
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
         cBack           =   16777152
      End
      Begin lvButton.lvButtons_H CmdKey1 
         Height          =   555
         Index           =   9
         Left            =   3915
         TabIndex        =   17
         Top             =   2025
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         Caption         =   "9"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
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
         cBack           =   16777152
      End
      Begin lvButton.lvButtons_H CmdKey1 
         Height          =   555
         Index           =   0
         Left            =   4725
         TabIndex        =   18
         ToolTipText     =   "Proses Settings"
         Top             =   2025
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
         Caption         =   "0"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
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
         cBack           =   16777152
      End
      Begin PRD.Liner Liner1 
         Height          =   30
         Left            =   135
         TabIndex        =   21
         Top             =   720
         Width           =   8430
         _ExtentX        =   14870
         _ExtentY        =   53
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCT :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   180
         TabIndex        =   20
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label lblProd 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Index           =   0
         Left            =   1395
         TabIndex        =   19
         Top             =   3780
         Width           =   2760
      End
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   6990
      Left            =   8910
      TabIndex        =   81
      Top             =   675
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   12330
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "i16x16"
      SmallIcons      =   "i16x16"
      ColHdrIcons     =   "i16x16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   8910
      TabIndex        =   80
      Top             =   180
      Width           =   1095
   End
End
Attribute VB_Name = "frmInputResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim srcItem                         As ListItem
Dim srcRecord                       As String
Dim prod_shift                      As Integer
Dim p_barang                        As String
Dim sSQL                            As String
Dim KeyFocus                        As Integer


Private Sub cboListProduct_Click()
    If cboListProduct.ListIndex = 0 Then
        FillListview p_eng_product_1, lvList
    ElseIf cboListProduct.ListIndex = 1 Then
        FillListview p_eng_product_2, lvList
    ElseIf cboListProduct.ListIndex = 2 Then
        FillListview p_eng_product_3, lvList
    ElseIf cboListProduct.ListIndex = 3 Then
        FillListview p_eng_product_4, lvList
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub AddSPB2(iQty As Variant, eng_prod As Variant)
On Error GoTo ErrHandler
 
    sSQL = "Insert Into sip_production.prod_result_logs (plant_mark,prod_machine_id,"
    sSQL = sSQL & " mkt_customer_id,eng_product_id,date,shift,product_status,qty,created_at,created_by,period_shift)"
    sSQL = sSQL & " values ('" & p_plant_mark & "','" & p_prod_machine_id & "'"
    sSQL = sSQL & " ,'" & p_mkt_customer_id & "','" & eng_prod & "','" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "'"
    sSQL = sSQL & " ,'" & prod_shift & "','" & p_barang & "','" & iQty & "'"
    sSQL = sSQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER.KODEUSER & "','" & Format(DTDate, "yyyy-mm-dd") & "')"
    
    sSQL_Insert sSQL

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
End Sub



Private Sub cmdKey1_Click(Index As Integer)
On Error GoTo ErrHandler
    
    If Index = 11 Then
        txtInput(0).Text = ""

    ElseIf Index = 10 Then
        If txtInput(0).Text <> "" Then

            If cboProduct.ListIndex = 0 Then
                AddSPB2 txtInput(0).Text, p_eng_product_1
                cboListProduct.ListIndex = 0
                FillListview p_eng_product_1, lvList
            
            ElseIf cboProduct.ListIndex = 1 Then
                AddSPB2 txtInput(0).Text, p_eng_product_2
                cboListProduct.ListIndex = 1
                FillListview p_eng_product_2, lvList
                
            ElseIf cboProduct.ListIndex = 2 Then
                AddSPB2 txtInput(0).Text, p_eng_product_3
                cboListProduct.ListIndex = 2
                FillListview p_eng_product_3, lvList
                
            ElseIf cboProduct.ListIndex = 3 Then
                AddSPB2 txtInput(0).Text, p_eng_product_4
                cboListProduct.ListIndex = 3
                FillListview p_eng_product_4, lvList
                
            End If

            txtInput(0).Text = ""
            
        End If
    Else
        txtInput(0).Text = txtInput(0).Text & CmdKey1(Index).Caption
    End If

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
    
End Sub


Private Sub CmdKeyBoard_Click(Index As Integer)
If Index = 0 Then
    KeyFocus = 0
    Picture1.Visible = True
    txtMBarcode_1.Text = ""
    txtMBarcode_1.SetFocus
ElseIf Index = 1 Then
    KeyFocus = 1

    Picture1.Visible = True
    txtMBarcode_2.Text = ""
    txtMBarcode_2.SetFocus
Else
    Picture1.Visible = False
End If

End Sub

Private Sub cmdNumber_Click(Index As Integer)
On Error GoTo ErrHandler


If KeyFocus = 0 Then

    If Index = 11 Then
        txtMBarcode_1.Text = ""
    ElseIf Index = 10 Then
        If txtMBarcode_1.Text = "" Then
            Picture1.Visible = False
        Else
            Picture1.Visible = False
        End If
    Else
        txtMBarcode_1.Text = txtMBarcode_1 & cmdNumber(Index).Caption
    End If
    
ElseIf KeyFocus = 1 Then

    If Index = 11 Then
        txtMBarcode_2.Text = ""
    ElseIf Index = 10 Then
        If txtMBarcode_2.Text = "" Then
            Picture1.Visible = False
        Else
            ProsesQuery txtMBarcode_1.Text & "-" & txtMBarcode_2.Text
            Picture1.Visible = False
        End If
    Else
        txtMBarcode_2.Text = txtMBarcode_2 & cmdNumber(Index).Caption
    End If
    
End If


Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
    
    
End Sub

Private Sub cmdOk_Click()
On Error GoTo ErrHandler

    Dim output()            As String
    Dim sSQL                As String
    Dim i                   As Integer

    output = Split(lblBarcode.Caption, "-")
    
    If txtEntry(3) <> p_customer_name Then
        MsgBox "Data Tidak Sesuai Customer", vbExclamation
        Exit Sub
    ElseIf Val(txtEntry(2).Text) <> p_machine_no Then
        MsgBox "Data Tidak Sesuai No Mesin..!", vbExclamation
        Exit Sub
    Else

        sSQL = "Insert Into sip_production.prod_result_logs (plant_mark,qc_label_product_id,box_number, prod_machine_id,"
        sSQL = sSQL & " mkt_customer_id,eng_product_id,date,shift,product_status,qty,created_at,created_by,period_shift)"
        sSQL = sSQL & " values ('" & p_plant_mark & "','" & output(0) & "','" & output(1) & "','" & p_prod_machine_id & "'"
        sSQL = sSQL & " ,'" & p_mkt_customer_id & "','" & lbleng_product_id.Caption & "','" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "'"
        sSQL = sSQL & " ,'" & prod_shift & "','" & p_barang & "','" & txtEntry(9).Text & "'"
        sSQL = sSQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER.KODEUSER & "','" & Format(DTDate, "yyyy-mm-dd") & "')"
        
        sSQL_Insert sSQL

    End If


    For i = 0 To 10
        txtEntry(i).Text = ""
    Next i
   
    txtBarcode.Text = ""
    lblBarcode.Caption = ""
    lbleng_product_id.Caption = ""
    cmdOk.Enabled = False
    txtBarcode.SetFocus

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
    
       
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
    Case vbKeyEscape
        Unload Me
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 27 Then
        Unload Me
    ElseIf KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    
    DTDate.Value = Format(p_shift, "yyyy-mm-dd")
    
    If Format(Now, "HH") >= 8 And Format(Now, "HH") <= 15 Then
        prod_shift = 1
        optshift1(0).Value = True
    ElseIf Format(Now, "HH") >= 16 And Format(Now, "HH") <= 23 Then
        prod_shift = 2
        optshift1(1).Value = True
    ElseIf Format(Now, "HH") >= 0 And Format(Now, "HH") <= 7 Then
        prod_shift = 3
        optshift1(2).Value = True
    End If
    
    optbarang1(0).Value = True

    If p_status_prod_1 = True Then
        cboProduct.AddItem p_prod_name_1
        cboListProduct.AddItem p_prod_name_1
    End If
   
    If p_status_prod_2 = True Then
        cboProduct.AddItem p_prod_name_2
        cboListProduct.AddItem p_prod_name_2
    End If

    If p_status_prod_3 = True Then
        FillListview p_eng_product_3, lvList
        cboProduct.AddItem p_prod_name_3
        cboListProduct.AddItem p_prod_name_3
    End If

    If p_status_prod_4 = True Then
        cboProduct.AddItem p_prod_name_4
        cboListProduct.AddItem p_prod_name_4
    End If
    
    cboListProduct.ListIndex = 0

    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
    
       
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmInputResult = Nothing
End Sub

Private Sub FillListview(eng_prod As String, listview As listview)
On Error GoTo ErrHandler

    Dim i As Integer
    Dim Rs As New Recordset
    Dim sSQL As String
    i = 1
    Rs.CursorLocation = adUseClient
    
    sSQL = "SELECT a.id, a.plant_mark,a.period_Shift,a.shift,a.product_status,b.number, "
    sSQL = sSQL & " d.internal_part_id,"
    sSQL = sSQL & " a.qty,a.qc_label_product_id,a.box_number FROM sip_production.prod_result_logs a "
    sSQL = sSQL & " LEFT JOIN sip_production.prod_machines as b  on a.prod_machine_id = b.id"
    sSQL = sSQL & " LEFT JOIN sip_production.eng_products as d  on a.eng_product_id = d.id"
    sSQL = sSQL & " WHERE "
    sSQL = sSQL & " a.plant_mark = '" & p_plant_mark & "' "
    sSQL = sSQL & " and a.prod_machine_id = '" & p_prod_machine_id & "'"
    sSQL = sSQL & " and a.mkt_customer_id = '" & p_mkt_customer_id & "'"
    sSQL = sSQL & " and a.eng_product_id = '" & eng_prod & "'"
    sSQL = sSQL & " and date(a.period_Shift) = '" & Format(DTDate, "yyyy-mm-dd") & "' order by a.id desc"
    
    Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
    
With listview
    
    .GridLines = True
    .View = lvwReport

    
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "NO."
    .ColumnHeaders.Add , , "DATE"
    .ColumnHeaders.Add , , "SHIFT"
    .ColumnHeaders.Add , , "BARCODE"
    .ColumnHeaders.Add , , "QTY"
    .ColumnHeaders.Add , , "STATUS"
    .ColumnHeaders.Add , , "INTERNAL PART"
    .ColumnHeaders.Add , , "ID"
    .ListItems.Clear
    
    i = 1
    Do While Not Rs.EOF
    Set srcItem = .ListItems.Add(, , i, 1, 1)
        srcItem.SubItems(1) = Format(Rs.Fields("period_Shift"), "yyyy-mm-dd")
        srcItem.SubItems(2) = Rs.Fields("shift")
        srcItem.SubItems(3) = Rs.Fields("qc_label_product_id") & "-" & Rs.Fields("box_number")
        srcItem.SubItems(4) = Rs.Fields("qty")
        srcItem.SubItems(5) = Rs.Fields("product_status")
        srcItem.SubItems(6) = Rs.Fields("internal_part_id")
        srcItem.SubItems(7) = IIf(IsNull(Rs.Fields("id")), "", Rs.Fields("id"))

        Rs.MoveNext
        i = i + 1
    Loop
    
    .SortOrder = lvwDescending
        
End With
Call lvSizeColumns(listview)

'AltLVBackground listview, vbWhite, &H8000000F, frmInputResult

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
 
 
End Sub

Private Sub LvList_DblClick()
    If MsgBox("Apakah anda yakin akan mengapus " & lvList.SelectedItem.SubItems(2) & " ?", vbQuestion + vbYesNo) = vbYes Then
        sSQL_Delete "DELETE FROM sip_production.prod_result_logs WHERE id = " & lvList.SelectedItem.SubItems(7) & ""
        Call cboListProduct_Click
    End If
End Sub



Private Sub optbarang1_Click(Index As Integer)
    If Index = 0 Then
        p_barang = "ok"
        picBarcode.Visible = True
    ElseIf Index = 1 Then
        p_barang = "sisa"
        picBarcode.Visible = False
    ElseIf Index = 2 Then
        p_barang = "hold"
        picBarcode.Visible = False
    End If
    
End Sub

Private Sub optshift1_Click(Index As Integer)
    If Index = 0 Then
        prod_shift = 1
    ElseIf Index = 1 Then
        prod_shift = 2
    ElseIf Index = 2 Then
        prod_shift = 3
    End If
End Sub




Private Sub txtBarcode_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub 'allow user to correct mistakes
    If KeyAscii >= 48 And KeyAscii <= 57 Then 'numbers 0 to 9
    ElseIf KeyAscii = 45 Or KeyAscii = 43 Then '+ and - keys
    ElseIf KeyAscii = 13# Then
        ProsesQuery txtBarcode.Text
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub ProsesQuery(ByVal sBarcode As String)
On Error GoTo ErrHandler

Dim Rs As New Recordset
Dim Rs_search As New Recordset
Dim sSQL As String
Dim i As Integer
Dim barcode() As String

    
    If InStr(sBarcode, "-") <> "0" Then
        barcode = Split(sBarcode, "-")
        
        Rs_search.CursorLocation = adUseClient
        
        Rs_search.Open "select * from sip_production.prod_result_logs a where a.qc_label_product_id = '" & barcode(0) & "' and " & _
                        "a.box_number = '" & barcode(1) & "'", CN, adOpenStatic, adLockOptimistic
        If Rs_search.RecordCount > 0 Then
            MsgBox "Data Sudah Pernah di Scan..!", vbExclamation
            For i = 0 To 10
                txtEntry(i).Text = ""
            Next i
            txtBarcode.Text = ""
            lblBarcode.Caption = ""
            lbleng_product_id.Caption = ""
            cmdOk.Enabled = False
            txtBarcode.SetFocus
            Exit Sub
        Else
            Rs.CursorLocation = adUseClient
            sSQL = "select a.id,a.seq,a.cavity,a.engine_number,a.quantity,a.quantity_box,a.eng_product_id,"
            sSQL = sSQL & " b.name as plant_name,c.name as customer_name,d.shift,e.nik,e.name as employee_name,"
            sSQL = sSQL & " f.internal_part_id,f.name as product_name, f.customer_part_number,f.customer_part_name,"
            sSQL = sSQL & " f.unix_code,f.model from sip_production.qc_label_products a"
            sSQL = sSQL & " left join sip_production.sys_plants b on a.sys_plant_id = b.id"
            sSQL = sSQL & " left join sip_production.mkt_customers c on a.mkt_customer_id = c.id"
            sSQL = sSQL & " left join sip_production.hrd_work_shifts d on a.hrd_work_shift_id = d.id"
            sSQL = sSQL & " left join sip_production.hrd_employees e on a.hrd_employee_id = e.id"
            sSQL = sSQL & " left join sip_production.eng_products f on a.eng_product_id = f.id"
            sSQL = sSQL & " where a.id = '" & barcode(0) & "'"
            
            Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
                If Rs.RecordCount < 1 Then
                    MsgBox "Data Tidak ditemukan", vbExclamation
                    For i = 0 To 10
                        txtEntry(i).Text = ""
                    Next i
                    txtBarcode.Text = ""
                    lblBarcode.Caption = ""
                    lbleng_product_id.Caption = ""
                    cmdOk.Enabled = False
                    txtBarcode.SetFocus
                    Exit Sub
                Else
                    cmdOk.Enabled = True
                    txtEntry(0).Text = Rs.Fields("product_name")
                    txtEntry(1).Text = Rs.Fields("customer_part_name")
                    'txtEntry(2).Text = Rs.Fields("internal_part_id")
                    'txtEntry(3).Text = Rs.Fields("customer_part_number")
                    txtEntry(4).Text = IIf(IsNull(Rs.Fields("unix_code")), "", Rs.Fields("unix_code"))
                    txtEntry(5).Text = IIf(IsNull(Rs.Fields("model")), "", Rs.Fields("model"))
                    txtEntry(6).Text = Rs.Fields("shift")
                    txtEntry(7).Text = IIf(IsNull(Rs.Fields("employee_name")), "", Rs.Fields("employee_name"))
                    txtEntry(8).Text = Rs.Fields("quantity")
                    txtEntry(9).Text = Rs.Fields("quantity_box")
                    txtEntry(10).Text = Rs.Fields("cavity")
                    txtEntry(2).Text = Rs.Fields("engine_number")
                    txtEntry(3).Text = Rs.Fields("customer_name")
                    lblBarcode.Caption = sBarcode
                    lbleng_product_id.Caption = Rs.Fields("eng_product_id")
                    sBarcode = ""
                    
                    cmdOk.SetFocus
                End If
        End If
        Set Rs = Nothing
        Set Rs_search = Nothing
    End If

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
     
     
End Sub


Private Sub txtBarcode_LostFocus()
    If txtBarcode.Text <> "" Then
        ProsesQuery txtBarcode.Text
    End If
End Sub


