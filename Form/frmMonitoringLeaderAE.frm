VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LvButton.ocx"
Begin VB.Form frmMonitoringLeaderAE 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Monitoring Leader"
   ClientHeight    =   6090
   ClientLeft      =   3600
   ClientTop       =   2265
   ClientWidth     =   9930
   Icon            =   "frmMonitoringLeaderAE.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9225
      Top             =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2850
      Left            =   4410
      ScaleHeight     =   2820
      ScaleWidth      =   5340
      TabIndex        =   0
      Top             =   1260
      Visible         =   0   'False
      Width           =   5370
      Begin lvButton.lvButtons_H cmdNumber 
         Height          =   735
         Index           =   13
         Left            =   2925
         TabIndex        =   14
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
         TabIndex        =   1
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
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   10
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
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   13
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
         TabIndex        =   15
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
      TabIndex        =   44
      Top             =   540
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   900
      Caption         =   "START CHECK"
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
   Begin VB.Frame frameLogin 
      BackColor       =   &H00FFFFFF&
      Height          =   1995
      Left            =   135
      TabIndex        =   35
      Top             =   1170
      Visible         =   0   'False
      Width           =   4245
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
         TabIndex        =   37
         Top             =   270
         Width           =   2040
      End
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
         TabIndex        =   36
         Top             =   810
         Width           =   2040
      End
      Begin lvButton.lvButtons_H cmdKeyUser 
         Height          =   420
         Left            =   3435
         TabIndex        =   38
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
         TabIndex        =   39
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
         Image           =   "frmMonitoringLeaderAE.frx":617A
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   480
         Left            =   2775
         TabIndex        =   40
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
         Image           =   "frmMonitoringLeaderAE.frx":C9DC
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdKeyPass 
         Height          =   420
         Left            =   3435
         TabIndex        =   41
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
         TabIndex        =   43
         Top             =   900
         Width           =   690
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
         TabIndex        =   42
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.Frame frameCheck 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Input Data"
      Enabled         =   0   'False
      Height          =   4200
      Left            =   135
      TabIndex        =   19
      Top             =   1170
      Width           =   9645
      Begin PRD.Liner Liner2 
         Height          =   30
         Left            =   90
         TabIndex        =   45
         Top             =   2925
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   53
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CHECK KESESUAIAN"
         Height          =   780
         Left            =   225
         TabIndex        =   32
         Top             =   315
         Width           =   3120
         Begin VB.OptionButton optKesesuaian 
            BackColor       =   &H00FFFFFF&
            Caption         =   "NO"
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
            Index           =   1
            Left            =   1575
            TabIndex        =   34
            Top             =   315
            Width           =   1230
         End
         Begin VB.OptionButton optKesesuaian 
            BackColor       =   &H00FFFFFF&
            Caption         =   "YES "
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
            Index           =   0
            Left            =   225
            TabIndex        =   33
            Top             =   315
            Value           =   -1  'True
            Width           =   1230
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CHECK MATERIAL"
         Height          =   780
         Left            =   225
         TabIndex        =   29
         Top             =   1170
         Width           =   3120
         Begin VB.OptionButton optMaterial 
            BackColor       =   &H00FFFFFF&
            Caption         =   "OK"
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
            Index           =   0
            Left            =   225
            TabIndex        =   31
            Top             =   315
            Value           =   -1  'True
            Width           =   1230
         End
         Begin VB.OptionButton optMaterial 
            BackColor       =   &H00FFFFFF&
            Caption         =   "NG"
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
            Index           =   1
            Left            =   1575
            TabIndex        =   30
            Top             =   315
            Width           =   1230
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CHECK ABNORMALITY"
         Height          =   780
         Left            =   225
         TabIndex        =   26
         Top             =   2025
         Width           =   3120
         Begin VB.OptionButton optAbnormality 
            BackColor       =   &H00FFFFFF&
            Caption         =   "TIDAK"
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
            Index           =   1
            Left            =   1575
            TabIndex        =   28
            Top             =   315
            Width           =   1230
         End
         Begin VB.OptionButton optAbnormality 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ADA"
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
            Index           =   0
            Left            =   225
            TabIndex        =   27
            Top             =   315
            Value           =   -1  'True
            Width           =   1230
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TARGET YIELD"
         Height          =   780
         Left            =   3870
         TabIndex        =   23
         Top             =   315
         Width           =   3120
         Begin VB.OptionButton optTarget 
            BackColor       =   &H00FFFFFF&
            Caption         =   "TERCAPAI"
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
            Index           =   0
            Left            =   225
            TabIndex        =   25
            Top             =   315
            Value           =   -1  'True
            Width           =   1230
         End
         Begin VB.OptionButton optTarget 
            BackColor       =   &H00FFFFFF&
            Caption         =   "TIDAK"
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
            Index           =   1
            Left            =   1575
            TabIndex        =   24
            Top             =   315
            Width           =   1230
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CYCLE TIME"
         Height          =   780
         Left            =   3870
         TabIndex        =   20
         Top             =   1170
         Width           =   3120
         Begin VB.OptionButton optCycleTime 
            BackColor       =   &H00FFFFFF&
            Caption         =   "TIDAK"
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
            Index           =   1
            Left            =   1575
            TabIndex        =   22
            Top             =   315
            Width           =   1230
         End
         Begin VB.OptionButton optCycleTime 
            BackColor       =   &H00FFFFFF&
            Caption         =   "TERCAPAI"
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
            Index           =   0
            Left            =   225
            TabIndex        =   21
            Top             =   315
            Value           =   -1  'True
            Width           =   1230
         End
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1665
         TabIndex        =   54
         Top             =   3825
         Width           =   195
      End
      Begin VB.Label lblLossTime 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
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
         Left            =   1980
         TabIndex        =   53
         Top             =   3825
         Width           =   3975
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "TIME CHECK"
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
         Left            =   270
         TabIndex        =   52
         Top             =   3825
         Width           =   1230
      End
      Begin VB.Label lblTimeStop 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1980
         TabIndex        =   51
         Top             =   3465
         Width           =   3975
      End
      Begin VB.Label lblTimeStart 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1980
         TabIndex        =   50
         Top             =   3105
         Width           =   3975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1665
         TabIndex        =   49
         Top             =   3465
         Width           =   195
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
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
         Left            =   1665
         TabIndex        =   48
         Top             =   3105
         Width           =   150
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "STOP"
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
         Left            =   270
         TabIndex        =   47
         Top             =   3465
         Width           =   1185
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "START"
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
         Left            =   270
         TabIndex        =   46
         Top             =   3105
         Width           =   1230
      End
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   9930
      TabIndex        =   16
      Top             =   0
      Width           =   9930
      Begin PRD.Liner Liner1 
         Height          =   30
         Left            =   0
         TabIndex        =   17
         Top             =   960
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   53
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "M O N I T OR I N G  S Y S T E M ®™"
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
         TabIndex        =   18
         Top             =   90
         Width           =   5355
      End
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   510
      Left            =   8100
      TabIndex        =   55
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
   Begin lvButton.lvButtons_H cmdAddNG 
      Height          =   555
      Left            =   8100
      TabIndex        =   56
      Top             =   5400
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
      Image           =   "frmMonitoringLeaderAE.frx":FC3E
      cBack           =   4210752
   End
End
Attribute VB_Name = "frmMonitoringLeaderAE"
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

Private Sub cmdAddNG_Click()
    frmNg.Show 1
End Sub

Private Sub cmdCancel_Click()
    frameLogin.Visible = False
    Picture1.Visible = False
End Sub

Private Sub cmdExit_Click()
    If CmdStart.Caption = "STOP CHECK" Then
        MsgBox "Silahkan STOP check jika mau Exit..!!", vbExclamation
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

Set RS_MONITORING = New ADODB.Recordset
If RS_MONITORING.State = adStateOpen Then RS_MONITORING.Close
RS_MONITORING.Open sSQL, CN, adOpenStatic, adLockReadOnly

If RS_MONITORING.BOF Or RS_MONITORING.EOF = True Then
    MsgBox "Username atau Password Salah, Silahkan Coba Lagi!", vbExclamation
    Exit Sub

ElseIf RS_MONITORING.Fields("status") = "suspend" Then
    MsgBox "User account is no longer active.Contact your administrator to re-activate your account!", vbExclamation
    Exit Sub

ElseIf Not RS_MONITORING.Fields("user") = txtUsername.text Then
    MsgBox "Username atau Password Salah, Silahkan Coba Lagi!", vbExclamation
    Exit Sub

ElseIf Not RS_MONITORING.Fields("password_clear") = txtPassword.text Then
    MsgBox "Username atau Password Salah, Silahkan Coba Lagi!", vbExclamation
    Exit Sub

Else

    If CmdStart.Caption = "START CHECK" Then
        frameCheck.Enabled = True
        CmdStart.Caption = "STOP CHECK"

        frameLogin.Visible = False
        Picture1.Visible = False
        
        Timer1.Enabled = True
        lblTimeStart.Caption = Format(Now, "yyyy-mm-dd hh:mm:ss")
        
        txtUsername.text = ""
        txtPassword.text = ""
        
    Else
        
        Timer1.Enabled = False
        
        Call UpdateData
        
        frameCheck.Enabled = False
        CmdStart.Caption = "START CHECK"
        frameLogin.Visible = False
        Picture1.Visible = False
    End If


End If

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
    
End Sub


Private Sub Timer1_Timer()
    lblTimeStop.Caption = Format(Now, "yyyy-mm-dd hh:mm:ss")
    lblLossTime.Caption = Format(CDate(CDate(lblTimeStop.Caption) - CDate(lblTimeStart.Caption)), "hh:mm:ss")

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
    frameLogin.Visible = True
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
CenterForm frmMonitoringLeader

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
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
    frmMonitoringLeader.CommandPass "Refresh"
    Set frmMonitoringLeader = Nothing
    Set RS_MONITORING = Nothing
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
    Dim check_kesesuaian As String
    Dim check_material As String
    Dim check_abnormality As String
    Dim target_yield As String
    Dim cycle_time As String
    
    If optKesesuaian(0).Value = True Then
        check_kesesuaian = "Y"
    Else
        check_kesesuaian = "N"
    End If
    
    If optMaterial(0).Value = True Then
        check_material = optMaterial(0).Caption
    Else
        check_material = optMaterial(1).Caption
    End If
    
    If optAbnormality(0).Value = True Then
        check_abnormality = optAbnormality(0).Caption
    Else
        check_abnormality = optAbnormality(1).Caption
    End If
    
    If optTarget(0).Value = True Then
        target_yield = optTarget(0).Caption
    Else
        target_yield = optTarget(1).Caption
    End If
    
    If optCycleTime(0).Value = True Then
        cycle_time = optCycleTime(0).Caption
    Else
        cycle_time = optCycleTime(1).Caption
    End If
    
    Dim iSQL As String
    iSQL = "INSERT INTO sip_production.prod_monitoring_leaders (plant_mark, "
    iSQL = iSQL & " prod_machine_id, "
    iSQL = iSQL & " mkt_customer_id , "
    iSQL = iSQL & " eng_product_id, "
    iSQL = iSQL & " Date, "
    iSQL = iSQL & " period_shift, "
    iSQL = iSQL & " hrd_employee_id, "
    iSQL = iSQL & " created_at, "
    iSQL = iSQL & " created_by, "
    iSQL = iSQL & " check_kesesuaian,"
    iSQL = iSQL & " check_material,"
    iSQL = iSQL & " check_abnormality,"
    iSQL = iSQL & " target_yield,"
    iSQL = iSQL & " cycle_time,"
    iSQL = iSQL & " start_check,"
    iSQL = iSQL & " stop_check,"
    iSQL = iSQL & " time_check) VALUES "
    iSQL = iSQL & " ('" & p_plant_mark & "', "
    iSQL = iSQL & " '" & p_prod_machine_id & "',"
    iSQL = iSQL & " '" & p_mkt_customer_id & "',"
    iSQL = iSQL & " '" & p_eng_product_1 & "',"
    iSQL = iSQL & " '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "',"
    iSQL = iSQL & " '" & Format(p_shift, "yyyy-mm-dd") & "',"
    iSQL = iSQL & " '" & RS_MONITORING.Fields("hrd_employee_id") & "',"
    iSQL = iSQL & " '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "',"
    iSQL = iSQL & " '" & RS_MONITORING.Fields("hrd_employee_id") & "',"
    iSQL = iSQL & " '" & check_kesesuaian & "',"
    iSQL = iSQL & " '" & check_material & "',"
    iSQL = iSQL & " '" & check_abnormality & "',"
    iSQL = iSQL & " '" & target_yield & "',"
    iSQL = iSQL & " '" & cycle_time & "',"
    iSQL = iSQL & " '" & lblTimeStart.Caption & "',"
    iSQL = iSQL & " '" & lblTimeStop.Caption & "',"
    iSQL = iSQL & " '" & lblLossTime.Caption & "')"

    sSQL_Insert iSQL
End Sub
