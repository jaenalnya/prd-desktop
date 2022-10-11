VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LvButton.ocx"
Begin VB.Form frmNg 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INPUT DATA PRODUCT NG (NOT GOOD)"
   ClientHeight    =   10125
   ClientLeft      =   4515
   ClientTop       =   2250
   ClientWidth     =   13710
   Icon            =   "frmNg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   13710
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H cboProduct_1 
      Height          =   690
      Left            =   180
      TabIndex        =   84
      Top             =   585
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   1217
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
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   16777152
   End
   Begin VB.Frame frameProd1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   8685
      TabIndex        =   80
      Top             =   5085
      Visible         =   0   'False
      Width           =   4245
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
         Index           =   0
         Left            =   45
         TabIndex        =   82
         Top             =   135
         Width           =   3480
      End
      Begin lvButton.lvButtons_H cmdOk1 
         Height          =   555
         Left            =   3600
         TabIndex        =   81
         Top             =   135
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   979
         Caption         =   "OK"
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "Keyboard"
      Height          =   3255
      Left            =   3825
      TabIndex        =   40
      Top             =   5895
      Visible         =   0   'False
      Width           =   8385
      Begin lvButton.lvButtons_H cmdKey 
         Height          =   600
         Index           =   0
         Left            =   585
         TabIndex        =   41
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
         TabIndex        =   42
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
         TabIndex        =   43
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
         TabIndex        =   44
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
         TabIndex        =   45
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
         TabIndex        =   46
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
         TabIndex        =   47
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
         TabIndex        =   48
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
         TabIndex        =   49
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
         TabIndex        =   50
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
         TabIndex        =   51
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
         TabIndex        =   52
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
         TabIndex        =   53
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
         TabIndex        =   54
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
         TabIndex        =   55
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
         TabIndex        =   56
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
         TabIndex        =   57
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
         TabIndex        =   58
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
         TabIndex        =   59
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
         TabIndex        =   60
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
         TabIndex        =   61
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
         TabIndex        =   62
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
         TabIndex        =   63
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
         TabIndex        =   64
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
         TabIndex        =   65
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
         TabIndex        =   66
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
         Left            =   2385
         TabIndex        =   67
         Top             =   2475
         Width           =   2535
         _ExtentX        =   4471
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
         Left            =   5040
         TabIndex        =   68
         Top             =   2475
         Width           =   1500
         _ExtentX        =   2646
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
         TabIndex        =   69
         Top             =   2475
         Width           =   1275
         _ExtentX        =   2249
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
         Left            =   6705
         TabIndex        =   70
         Top             =   1755
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1058
         Caption         =   "<--  Backspace"
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
         Left            =   1575
         TabIndex        =   71
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
      Begin lvButton.lvButtons_H cmdHide 
         Height          =   510
         Left            =   7020
         TabIndex        =   83
         Top             =   2520
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   900
         Caption         =   "HIDE"
         CapAlign        =   2
         BackStyle       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
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
   End
   Begin PRD.InactiveTimer itmrClose 
      Left            =   13050
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      Enabled         =   0   'False
   End
   Begin lvButton.lvButtons_H cmdDel1 
      Height          =   375
      Left            =   13050
      TabIndex        =   33
      Top             =   5220
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   661
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
      ImgAlign        =   4
      Image           =   "frmNg.frx":617A
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5055
      Left            =   4680
      TabIndex        =   1
      Top             =   45
      Width           =   8880
      Begin PRD.Liner Liner4 
         Height          =   30
         Left            =   0
         TabIndex        =   29
         Top             =   495
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   53
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   585
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [0 1]     DITRY START"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   2
         Left            =   1620
         TabIndex        =   3
         Top             =   585
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [0 2]     DENTED"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         ImgAlign        =   4
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   3
         Left            =   3060
         TabIndex        =   4
         Top             =   585
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "     [0 3]      OIL MARK"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   4
         Left            =   4500
         TabIndex        =   5
         Top             =   585
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [0 4]    NEMPEL"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   5
         Left            =   5940
         TabIndex        =   6
         Top             =   585
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [0 5]    BENANG"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   6
         Left            =   7380
         TabIndex        =   7
         Top             =   585
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "     [0 6]     NG INSERT"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   7
         Left            =   180
         TabIndex        =   8
         Top             =   1215
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [0 7]    BURN MARK"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   8
         Left            =   1620
         TabIndex        =   9
         Top             =   1215
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [0 8]     SILVER"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   9
         Left            =   3060
         TabIndex        =   10
         Top             =   1215
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [0 9]     SHOT MOLD"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   10
         Left            =   4500
         TabIndex        =   11
         Top             =   1215
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [1 0]     BLACK DOT"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   11
         Left            =   5940
         TabIndex        =   12
         ToolTipText     =   "WARNA TIDAK CAMPUR"
         Top             =   1215
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [1 1]   WRN T. CAMP"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   12
         Left            =   7380
         TabIndex        =   13
         Top             =   1215
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [1 2]     WARNA BEDA"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   13
         Left            =   180
         TabIndex        =   14
         Top             =   1845
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [1 3]    SCRATCH"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   14
         Left            =   1620
         TabIndex        =   15
         Top             =   1845
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [1 4]     PECAH"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   15
         Left            =   3060
         TabIndex        =   16
         Top             =   1845
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [1 5]     DIMENSI OUT"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   16
         Left            =   4500
         TabIndex        =   17
         ToolTipText     =   "Proses Settings"
         Top             =   1845
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [1 6]     PROSES SET"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   17
         Left            =   5940
         TabIndex        =   18
         Top             =   1845
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [1 7]     WELD LINE"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   18
         Left            =   7380
         TabIndex        =   19
         Top             =   1845
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [1 8]      CLOUDY-ASAP"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   19
         Left            =   180
         TabIndex        =   20
         ToolTipText     =   "Silver After Lap"
         Top             =   2475
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [1 9]    SILVER A.LAP"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   20
         Left            =   1620
         TabIndex        =   21
         Top             =   2475
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [2 0]    FLOW MARK"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   21
         Left            =   3060
         TabIndex        =   22
         Top             =   2475
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [2 1]    WHITE DOT"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   22
         Left            =   4500
         TabIndex        =   23
         Top             =   2475
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [2 2]    BUBBLE"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   23
         Left            =   5940
         TabIndex        =   24
         Top             =   2475
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [2 3]    SHINNING"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   24
         Left            =   7380
         TabIndex        =   25
         ToolTipText     =   "Crack Cutting"
         Top             =   2475
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [2 4]    CRACK CUT"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   25
         Left            =   180
         TabIndex        =   30
         Top             =   3105
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [2 5]    COROSIF"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   26
         Left            =   1620
         TabIndex        =   31
         Top             =   3105
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [2 6]    GAS MARK"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   27
         Left            =   3060
         TabIndex        =   32
         Top             =   3105
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [2 7]     WHITE MARK"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   28
         Left            =   4500
         TabIndex        =   34
         Top             =   3105
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [2 8]     OVER CUT"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   29
         Left            =   5940
         TabIndex        =   35
         Top             =   3105
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [2 9]    SHINK MARK"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   30
         Left            =   7380
         TabIndex        =   36
         Top             =   3105
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [3 0]     OVER HEAT"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   31
         Left            =   180
         TabIndex        =   37
         Top             =   3735
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [3 1]   PROD JATUH"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   32
         Left            =   1620
         TabIndex        =   38
         Top             =   3735
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [3 2]    TEMBOLOK"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   33
         Left            =   5940
         TabIndex        =   39
         Top             =   4365
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [3 3]     OTHER"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   34
         Left            =   3060
         TabIndex        =   72
         Top             =   3735
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [3 4]     DIRTY"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   35
         Left            =   4500
         TabIndex        =   73
         Top             =   3735
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [3 5]     EJECTOR M"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   36
         Left            =   5940
         TabIndex        =   74
         Top             =   3735
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [3 6]  DISCOLOUR"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   37
         Left            =   7380
         TabIndex        =   75
         Top             =   3735
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [3 7]   SAMPLE QC"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   38
         Left            =   180
         TabIndex        =   76
         Top             =   4365
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [3 8]  BERCAK AIR"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   39
         Left            =   1620
         TabIndex        =   77
         Top             =   4365
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [3 9]   OVER PACK"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   40
         Left            =   3060
         TabIndex        =   78
         Top             =   4365
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [4 0]   GATE BOLONG"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin lvButton.lvButtons_H cmdNG1 
         Height          =   600
         Index           =   41
         Left            =   4500
         TabIndex        =   79
         Top             =   4365
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1058
         Caption         =   "    [4 1]  DOUBLE INJ"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
         cBack           =   14737632
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCT :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         TabIndex        =   28
         Top             =   135
         Width           =   1680
      End
      Begin VB.Label lblProd 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         Left            =   1620
         TabIndex        =   27
         Top             =   135
         Width           =   6495
      End
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   4245
      Left            =   180
      TabIndex        =   26
      Top             =   5130
      Width           =   12750
      _ExtentX        =   22490
      _ExtentY        =   7488
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
      BackColor       =   16777215
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
   Begin MSComctlLib.ImageList i16x16 
      Left            =   7560
      Top             =   810
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
            Picture         =   "frmNg.frx":1CB3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNg.frx":1D54E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNg.frx":1DF60
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNg.frx":1E2FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNg.frx":1E694
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNg.frx":1EA2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNg.frx":1EDC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNg.frx":1F7DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNg.frx":201EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNg.frx":20BFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNg.frx":21610
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNg.frx":22022
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNg.frx":22A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNg.frx":23446
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNg.frx":239E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   510
      Left            =   6300
      TabIndex        =   0
      Top             =   9495
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   900
      Caption         =   "E&XIT"
      CapAlign        =   2
      BackStyle       =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
   Begin lvButton.lvButtons_H cboProduct_2 
      Height          =   690
      Left            =   180
      TabIndex        =   86
      Top             =   1800
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   1217
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
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12648447
   End
   Begin lvButton.lvButtons_H cboProduct_3 
      Height          =   690
      Left            =   180
      TabIndex        =   87
      Top             =   3060
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   1217
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
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12640511
   End
   Begin lvButton.lvButtons_H cboProduct_4 
      Height          =   690
      Left            =   180
      TabIndex        =   91
      Top             =   4275
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   1217
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
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12648384
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT 4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   270
      TabIndex        =   90
      Top             =   3915
      Width           =   4155
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   270
      TabIndex        =   89
      Top             =   2700
      Width           =   4155
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   270
      TabIndex        =   88
      Top             =   1440
      Width           =   4155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   270
      TabIndex        =   85
      Top             =   225
      Width           =   4155
   End
End
Attribute VB_Name = "frmNg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim srcItem                         As ListItem
Dim srcRecord                       As String
Dim iKeyboard                       As Integer
Dim AutoClose                       As Boolean
Dim iProduct                        As String


Private Sub cboProduct_1_Click()
    FillListview p_eng_product_1, lvList
    LblProd(0).Caption = p_prod_name_1
    iProduct = p_eng_product_1

    Dim i As Integer
    i = 1
    For i = 1 To 41
        cmdNG1(i).BackColor = &HFFFFC0
    Next i
End Sub

Private Sub cboProduct_2_Click()
    FillListview p_eng_product_2, lvList
    LblProd(0).Caption = p_prod_name_2
    iProduct = p_eng_product_2
    Dim i As Integer
    i = 1
    For i = 1 To 41
        cmdNG1(i).BackColor = &HC0FFFF
    Next i
    
End Sub

Private Sub cboProduct_3_Click()
    FillListview p_eng_product_3, lvList
    LblProd(0).Caption = p_prod_name_3
    iProduct = p_eng_product_3

    Dim i As Integer
    i = 1
    For i = 1 To 41
        cmdNG1(i).BackColor = &HC0E0FF
    Next i
        
End Sub

Private Sub cboProduct_4_Click()
    FillListview p_eng_product_4, lvList
    LblProd(0).Caption = p_prod_name_4
    iProduct = p_eng_product_4

    Dim i As Integer
    i = 1
    For i = 1 To 41
        cmdNG1(i).BackColor = &HC0FFC0
    Next i
        
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Addng(i_ng As Variant, eng_prod As Variant, Optional sdesc As String)
On Error GoTo ErrHandler

    Dim sQL As String

    sQL = "insert into sip_production.prod_ng_logs (plant_mark,prod_machine_id,mkt_customer_id,eng_product_id,prod_ng_id,"
    sQL = sQL & " date,period_shift,description,created_at,created_by) values"
    sQL = sQL & " ('" & p_plant_mark & "','" & p_prod_machine_id & "','" & p_mkt_customer_id & "','" & eng_prod & "','" & i_ng & "'"
    sQL = sQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & Format(p_shift, "yyyy-mm-dd") & "','" & sdesc & "'"
    sQL = sQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER.KODEUSER & "')"
    
    sSQL_Insert sQL

    

    Dim Rs As New Recordset
    Dim strSQL As String
    Dim Counter As Variant

    Counter = 0
    Rs.CursorLocation = adUseClient '
    
    strSQL = "select a.* from sip_production.prod_data_ngs a where "
    strSQL = strSQL & " a.plant_mark = '" & p_plant_mark & "' "
    strSQL = strSQL & " and a.prod_machine_id = '" & p_prod_machine_id & "'"
    strSQL = strSQL & " and a.mkt_customer_id = '" & p_mkt_customer_id & "'"
    strSQL = strSQL & " and a.eng_product_id = '" & eng_prod & "'"
    strSQL = strSQL & " and a.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "'"
    strSQL = strSQL & " and a.period_hour = '" & Format(Now, "HH") & "'"
    strSQL = strSQL & " and a.prod_ng_id = '" & i_ng & "'"
    strSQL = strSQL & " and a.description = '" & sdesc & "'"
    
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open strSQL, CN, adOpenDynamic, adLockPessimistic
    
    If Rs.RecordCount < 1 Then

        strSQL = "insert into sip_production.prod_data_ngs (plant_mark,prod_machine_id,mkt_customer_id,eng_product_id,"
        strSQL = strSQL & " date,period_shift,period_hour,prod_ng_id,counter_ng,description,operator_1,operator_2,created_at,created_by) values"
        strSQL = strSQL & " ('" & p_plant_mark & "'"
        strSQL = strSQL & " ,'" & p_prod_machine_id & "'"
        strSQL = strSQL & " ,'" & p_mkt_customer_id & "'"
        strSQL = strSQL & " ,'" & eng_prod & "'"
        strSQL = strSQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "'"
        strSQL = strSQL & " ,'" & Format(p_shift, "yyyy-mm-dd") & "'"
        strSQL = strSQL & " ,'" & Format(Now, "HH") & "'"
        strSQL = strSQL & " ,'" & i_ng & "'"
        strSQL = strSQL & " ,'" & 1 & "'"
        strSQL = strSQL & " ,'" & sdesc & "'"
        strSQL = strSQL & " ,'" & ACTIVE_USER.KODEUSER & "'"
        strSQL = strSQL & " ,'" & ACTIVE_USER_2.KODEUSER & "'"
        strSQL = strSQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "'"
        strSQL = strSQL & " ,'" & ACTIVE_USER.KODEUSER & "')"
        
        sSQL_Insert strSQL
        
    Else
        Counter = Val(Rs.Fields("counter_ng")) + 1
        strSQL = "update sip_production.prod_data_ngs set "
            strSQL = strSQL & " counter_ng = '" & Counter & "'"
            strSQL = strSQL & " ,operator_1 = '" & ACTIVE_USER.KODEUSER & "'"
            strSQL = strSQL & " ,operator_2 = '" & ACTIVE_USER_2.KODEUSER & "'"
            strSQL = strSQL & " ,date = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' "
            strSQL = strSQL & " where "
            strSQL = strSQL & " plant_mark = '" & p_plant_mark & "' "
            strSQL = strSQL & " and prod_machine_id = '" & p_prod_machine_id & "'"
            strSQL = strSQL & " and mkt_customer_id = '" & p_mkt_customer_id & "'"
            strSQL = strSQL & " and eng_product_id = '" & eng_prod & "'"
            strSQL = strSQL & " and period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "'"
            strSQL = strSQL & " and period_hour = '" & Format(Now, "HH") & "'"
            strSQL = strSQL & " and prod_ng_id = '" & i_ng & "'"
            strSQL = strSQL & " and description = '" & sdesc & "'"
            
        sSQL_Update strSQL
        
        Set Rs = Nothing
     
    End If
    

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
End Sub


Private Sub cmdDel1_Click()
On Error GoTo ErrHandler

If MsgBox("Apakah akan menghapus data NG", vbQuestion + vbCritical + vbYesNo) = vbYes Then

    sSQL_Update "UPDATE sip_production.prod_ng_logs SET status = 'suspend' WHERE id = " & lvList.SelectedItem.SubItems(1) & ""
    FillListview iProduct, lvList
    
    
    Dim Rs As New Recordset
    Dim strSQL As String
    Dim Counter As Variant

    Counter = 0
    Rs.CursorLocation = adUseClient '
    
    strSQL = "select a.* from sip_production.prod_data_ngs a where "
    strSQL = strSQL & " a.plant_mark = '" & p_plant_mark & "' "
    strSQL = strSQL & " and a.prod_machine_id = '" & p_prod_machine_id & "'"
    strSQL = strSQL & " and a.mkt_customer_id = '" & p_mkt_customer_id & "'"
    strSQL = strSQL & " and a.eng_product_id = '" & iProduct & "'"
    strSQL = strSQL & " and a.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "'"
    strSQL = strSQL & " and a.period_hour = '" & Format(lvList.SelectedItem.SubItems(2), "HH") & "'"
    strSQL = strSQL & " and a.prod_ng_id = '" & lvList.SelectedItem.SubItems(3) & "'"
    strSQL = strSQL & " and a.description = '" & lvList.SelectedItem.SubItems(6) & "'"
    
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open strSQL, CN, adOpenDynamic, adLockPessimistic
    
    If Rs.RecordCount > 0 Then


        Counter = Val(Rs.Fields("counter_ng")) - 1
        strSQL = "update sip_production.prod_data_ngs set "
            strSQL = strSQL & " counter_ng = '" & Counter & "'"
            strSQL = strSQL & " ,operator_1 = '" & ACTIVE_USER.KODEUSER & "'"
            strSQL = strSQL & " ,operator_2 = '" & ACTIVE_USER_2.KODEUSER & "'"
            strSQL = strSQL & " ,date = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' "
            strSQL = strSQL & " where "
            strSQL = strSQL & " plant_mark = '" & p_plant_mark & "' "
            strSQL = strSQL & " and prod_machine_id = '" & p_prod_machine_id & "'"
            strSQL = strSQL & " and mkt_customer_id = '" & p_mkt_customer_id & "'"
            strSQL = strSQL & " and eng_product_id = '" & iProduct & "'"
            strSQL = strSQL & " and period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "'"
            strSQL = strSQL & " and period_hour = '" & Format(lvList.SelectedItem.SubItems(2), "HH") & "'"
            strSQL = strSQL & " and prod_ng_id = '" & lvList.SelectedItem.SubItems(3) & "'"
            strSQL = strSQL & " and description = '" & lvList.SelectedItem.SubItems(6) & "'"
            
        sSQL_Update strSQL
        
        Set Rs = Nothing
     
    End If
End If
    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
    
End Sub


Private Sub cmdHide_Click()
    Frame3.Visible = False
    iKeyboard = 0
    frameProd1.Visible = False
End Sub

Private Sub cmdKey_Click(Index As Integer)
If iKeyboard = 1 Then
    If Index = 27 Then
        If txtLain(0).text = "" Then
            MsgBox "Keterangan Belum di Isi..!", vbExclamation
            Exit Sub
        Else
            Frame3.Visible = False
            iKeyboard = 0
        End If
    ElseIf Index = 28 Then
        txtLain(0).text = ""
    ElseIf Index = 29 Then
        If Len(txtLain(0).text) = 0 Then Exit Sub
        txtLain(0).text = Mid(txtLain(0).text, 1, Len(txtLain(0).text) - 1)
    
    ElseIf Index = 26 Then
        txtLain(0).text = txtLain(0).text & " "
    Else
        txtLain(0).text = txtLain(0).text & cmdKey(Index).Caption
    End If
ElseIf iKeyboard = 2 Then
    If Index = 27 Then
        If txtLain(1).text = "" Then
            MsgBox "Keterangan Belum di Isi..!", vbExclamation
            Exit Sub
        Else
            Frame3.Visible = False
            iKeyboard = 0
        End If
    ElseIf Index = 28 Then
        txtLain(1).text = ""
    ElseIf Index = 29 Then
        If Len(txtLain(1).text) = 0 Then Exit Sub
        txtLain(1).text = Mid(txtLain(1).text, 1, Len(txtLain(1).text) - 1)
    ElseIf Index = 26 Then
        txtLain(1).text = txtLain(1).text & " "
    Else
        txtLain(1).text = txtLain(1).text & cmdKey(Index).Caption
    End If
End If

End Sub

Private Sub cmdNG1_Click(Index As Integer)
On Error GoTo ErrHandler

    If Index = 33 Then
        iKeyboard = 1
        Frame3.Visible = True
        frameProd1.Visible = True
        txtLain(0).SetFocus
    Else
        Addng Index, iProduct
        FillListview iProduct, lvList
        frameProd1.Visible = False
    End If

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
    
End Sub


Private Sub cmdOk1_Click()
    Addng 33, iProduct, txtLain(0).text
    FillListview iProduct, lvList
    frameProd1.Visible = False
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

    If p_status_prod_1 = True Then
        iProduct = p_eng_product_1
        FillListview iProduct, lvList
        LblProd(0).Caption = p_prod_name_1
        cboProduct_1.Caption = p_prod_name_1
        
        Dim i As Integer
        i = 1
        For i = 1 To 41
            cmdNG1(i).BackColor = &HFFFFC0
        Next i
    
    Else
        cboProduct_1.Enabled = False
        cboProduct_1.Caption = ""
    End If
   
    If p_status_prod_2 = True Then
        cboProduct_2.Enabled = True
        cboProduct_2.Caption = p_prod_name_2
    Else
        cboProduct_2.Enabled = False
        cboProduct_2.Caption = ""
    End If

    If p_status_prod_3 = True Then
        cboProduct_3.Enabled = True
        cboProduct_3.Caption = p_prod_name_3
    Else
        cboProduct_3.Enabled = False
        cboProduct_3.Caption = ""
    End If

    If p_status_prod_4 = True Then
        cboProduct_4.Enabled = True
        cboProduct_4.Caption = p_prod_name_4
    Else
        cboProduct_4.Enabled = False
        cboProduct_4.Caption = ""
    End If
    

    If ReadINI("SETTING", "NGAUTOCLOSE", App.Path & "\Settings.ini") <> "" Then
        AutoClose = ReadINI("SETTING", "NGAUTOCLOSE", App.Path & "\Settings.ini")
        If AutoClose = True Then
            itmrClose.InactiveInterval = 1000 * 10
            itmrClose.Enabled = True
        End If
    End If
    
    formNG = True
    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    formNG = False
    Set frmNg = Nothing
    Set RS_NG = Nothing
End Sub

Private Sub FillListview(eng_prod As String, listview As listview)
On Error GoTo ErrHandler

    Dim i As Integer
    Dim Rs As New Recordset
    Dim sSQL As String

    i = 1
    Rs.CursorLocation = adUseClient

    
    sSQL = "SELECT @no:=@no+1 AS nomor,a.id, a.plant_mark,a.date,a.prod_ng_id,b.name as ng_name, a.description,a.status "
    sSQL = sSQL & " FROM sip_production.prod_ng_logs a JOIN (SELECT @no:=0) as no"
    sSQL = sSQL & " INNER JOIN sip_production.prod_ngs as b  on a.prod_ng_id = b.id "
    sSQL = sSQL & " WHERE "
    sSQL = sSQL & " a.plant_mark = '" & p_plant_mark & "' "
    sSQL = sSQL & " and a.prod_machine_id = '" & p_prod_machine_id & "'"
    sSQL = sSQL & " and a.mkt_customer_id = '" & p_mkt_customer_id & "'"
    sSQL = sSQL & " and a.eng_product_id = '" & eng_prod & "'"
    sSQL = sSQL & " and a.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "'"
    sSQL = sSQL & " ORDER BY nomor DESC LIMIT 10 "
    
    Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
    
With listview
    
    .GridLines = True
    .View = lvwReport
    .SortOrder = lvwDescending
    
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "NO.", 600
    .ColumnHeaders.Add , , "ID", 800
    .ColumnHeaders.Add , , "TANGGAL", 2000
    .ColumnHeaders.Add , , "KODE NG", 1000
    .ColumnHeaders.Add , , "NAMA NG", 2100
    .ColumnHeaders.Add , , "STATUS", 1000
    .ColumnHeaders.Add , , "DESCRIPTION", 2000
    .ListItems.Clear
    Do While Not Rs.EOF
    Set srcItem = .ListItems.Add(, , Rs.Fields("nomor"), 1, 1)
        srcItem.SubItems(1) = Rs.Fields("id")
        srcItem.SubItems(2) = Format(Rs.Fields("date"), "yyyy-mm-dd hh:mm:ss")
        srcItem.SubItems(3) = Rs.Fields("prod_ng_id")
        srcItem.SubItems(4) = Rs.Fields("ng_name")
        srcItem.SubItems(5) = Rs.Fields("status")
        srcItem.SubItems(6) = IIf(IsNull(Rs.Fields("description")), "", Rs.Fields("description"))
        
        Rs.MoveNext
        
    Loop
        
End With
'Call lvSizeColumns(listview)

'AltLVBackground listview, vbWhite, &HFFFFC0, frmNg

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
 
 
End Sub

Private Sub itmrClose_UserInactive()
On Error Resume Next
    Unload Me
End Sub
