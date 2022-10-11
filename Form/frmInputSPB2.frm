VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmInputSPB2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12315
   Icon            =   "frmInputSPB2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   12315
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3660
      Left            =   45
      TabIndex        =   22
      Top             =   495
      Width           =   6045
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   315
         TabIndex        =   23
         Top             =   675
         Width           =   5370
      End
      Begin PRD.Liner Liner1 
         Height          =   30
         Left            =   90
         TabIndex        =   24
         Top             =   495
         Width           =   5820
         _ExtentX        =   10266
         _ExtentY        =   53
      End
      Begin lvButton.lvButtons_H cmdKey1 
         Height          =   555
         Index           =   1
         Left            =   765
         TabIndex        =   25
         Top             =   1440
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "1"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey1 
         Height          =   555
         Index           =   2
         Left            =   1935
         TabIndex        =   26
         Top             =   1440
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "2"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey1 
         Height          =   555
         Index           =   3
         Left            =   3105
         TabIndex        =   27
         Top             =   1440
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "3"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey1 
         Height          =   555
         Index           =   11
         Left            =   4275
         TabIndex        =   28
         Top             =   2160
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "CLEAR"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey1 
         Height          =   555
         Index           =   4
         Left            =   765
         TabIndex        =   29
         Top             =   2160
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "4"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey1 
         Height          =   555
         Index           =   5
         Left            =   1935
         TabIndex        =   30
         Top             =   2160
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "5"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey1 
         Height          =   555
         Index           =   6
         Left            =   3105
         TabIndex        =   31
         Top             =   2160
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "6"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey1 
         Height          =   555
         Index           =   10
         Left            =   4275
         TabIndex        =   32
         Top             =   1440
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "ENTER"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey1 
         Height          =   555
         Index           =   7
         Left            =   765
         TabIndex        =   33
         Top             =   2880
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "7"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey1 
         Height          =   555
         Index           =   8
         Left            =   1935
         TabIndex        =   34
         Top             =   2880
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "8"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey1 
         Height          =   555
         Index           =   9
         Left            =   3105
         TabIndex        =   35
         Top             =   2880
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "9"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey1 
         Height          =   555
         Index           =   0
         Left            =   4275
         TabIndex        =   36
         ToolTipText     =   "Proses Settings"
         Top             =   2880
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "0"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
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
         Height          =   330
         Index           =   0
         Left            =   1305
         TabIndex        =   38
         Top             =   135
         Width           =   4605
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
         TabIndex        =   37
         Top             =   135
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   3660
      Left            =   6165
      TabIndex        =   3
      Top             =   495
      Width           =   6045
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   315
         TabIndex        =   21
         Top             =   675
         Width           =   5370
      End
      Begin PRD.Liner Liner5 
         Height          =   30
         Left            =   90
         TabIndex        =   4
         Top             =   495
         Width           =   5820
         _ExtentX        =   10266
         _ExtentY        =   53
      End
      Begin lvButton.lvButtons_H cmdKey2 
         Height          =   555
         Index           =   1
         Left            =   765
         TabIndex        =   5
         Top             =   1440
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "1"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey2 
         Height          =   555
         Index           =   2
         Left            =   1935
         TabIndex        =   6
         Top             =   1440
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "2"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey2 
         Height          =   555
         Index           =   3
         Left            =   3105
         TabIndex        =   7
         Top             =   1440
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "3"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey2 
         Height          =   555
         Index           =   11
         Left            =   4275
         TabIndex        =   8
         Top             =   2160
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "CLEAR"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey2 
         Height          =   555
         Index           =   4
         Left            =   765
         TabIndex        =   9
         Top             =   2160
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "4"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey2 
         Height          =   555
         Index           =   5
         Left            =   1935
         TabIndex        =   10
         Top             =   2160
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "5"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey2 
         Height          =   555
         Index           =   6
         Left            =   3105
         TabIndex        =   11
         Top             =   2160
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "6"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey2 
         Height          =   555
         Index           =   10
         Left            =   4275
         TabIndex        =   12
         Top             =   1440
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "ENTER"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey2 
         Height          =   555
         Index           =   7
         Left            =   765
         TabIndex        =   13
         Top             =   2880
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "7"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey2 
         Height          =   555
         Index           =   8
         Left            =   1935
         TabIndex        =   14
         Top             =   2880
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "8"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey2 
         Height          =   555
         Index           =   9
         Left            =   3105
         TabIndex        =   15
         Top             =   2880
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "9"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H cmdKey2 
         Height          =   555
         Index           =   0
         Left            =   4275
         TabIndex        =   16
         ToolTipText     =   "Proses Settings"
         Top             =   2880
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   979
         Caption         =   "0"
         CapAlign        =   2
         BackStyle       =   6
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
         cBack           =   12648447
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
         Left            =   180
         TabIndex        =   18
         Top             =   135
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
         Height          =   330
         Index           =   1
         Left            =   1305
         TabIndex        =   17
         Top             =   135
         Width           =   4605
      End
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      ScaleHeight     =   450
      ScaleWidth      =   12315
      TabIndex        =   0
      Top             =   0
      Width           =   12315
      Begin lvButton.lvButtons_H cmdClose 
         Height          =   375
         Left            =   11070
         TabIndex        =   1
         Top             =   45
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   661
         Caption         =   "&Close"
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
         Image           =   "frmInputSPB2.frx":617A
         cBack           =   -2147483633
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "INPUT HASIL PRODUKSI (OK)"
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
         Height          =   255
         Left            =   4590
         TabIndex        =   2
         Top             =   90
         Width           =   3495
      End
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   4605
      Left            =   45
      TabIndex        =   19
      Top             =   4230
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   8123
      View            =   3
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView LvList2 
      Height          =   4605
      Left            =   6165
      TabIndex        =   20
      Top             =   4230
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   8123
      View            =   3
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
      Left            =   6120
      Top             =   3420
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
            Picture         =   "frmInputSPB2.frx":C304
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputSPB2.frx":CD16
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputSPB2.frx":D728
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputSPB2.frx":DAC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputSPB2.frx":DE5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputSPB2.frx":E1F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputSPB2.frx":E590
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputSPB2.frx":EFA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputSPB2.frx":F9B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputSPB2.frx":103C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputSPB2.frx":10DD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputSPB2.frx":117EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputSPB2.frx":121FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputSPB2.frx":12C0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputSPB2.frx":131AA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin PRD.InactiveTimer itmrClose 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "frmInputSPB2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim srcItem                         As ListItem
Dim srcRecord                       As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub AddSPB2(iQty As Integer, eng_prod As Variant)
    Dim sSQL As String
    Dim p_shift As Date

    If Format(Now, "HH") >= 0 And Format(Now, "HH") <= 7 Then
        p_shift = Format(DateAdd("d", -1, Format(Now, "yyyy-mm-dd")), "yyyy-mm-dd")
    Else
        p_shift = Format(Now, "yyyy-mm-dd")
    End If

    sSQL = "Insert Into sip_production.prod_spb2s_logs (plant_mark,qc_label_product_id,prod_machine_id,"
    sSQL = sSQL & " mkt_customer_id,eng_product_id,date,period_shift,qty,created_at,created_by)"
    sSQL = sSQL & " values ('" & p_plant_mark & "','0','" & p_prod_machine_id & "'"
    sSQL = sSQL & " ,'" & p_mkt_customer_id & "','" & eng_prod & "','" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "'"
    sSQL = sSQL & " ,'" & Format(p_shift, "yyyy-mm-dd") & "','" & iQty & "'"
    sSQL = sSQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER.KODEUSER & "')"
    
    sSQL_Insert sSQL

End Sub



Private Sub cmdKey1_Click(Index As Integer)
    If Index = 11 Then
        txtInput(0).Text = ""

    ElseIf Index = 10 Then
        AddSPB2 txtInput(0).Text, p_eng_product_1
        FillListview p_eng_product_1, lvList
        txtInput(0).Text = ""
    Else
        txtInput(0).Text = txtInput(0).Text & cmdKey1(Index).Caption
    End If
End Sub

Private Sub cmdKey2_Click(Index As Integer)
    If Index = 11 Then
        txtInput(1).Text = ""

    ElseIf Index = 10 Then
        AddSPB2 txtInput(1).Text, p_eng_product_2
        FillListview p_eng_product_2, LvList2
        txtInput(1).Text = ""
    Else
        txtInput(1).Text = txtInput(1).Text & cmdKey2(Index).Caption
    End If
End Sub

Private Sub Form_Activate()
    With MAIN
        Me.BackColor = .ACPMenu.BackColor
        Frame1.BackColor = .ACPMenu.BackColor
        Frame2.BackColor = .ACPMenu.BackColor
        Me.Picture = .ACPMenu.LoadBackground
    End With
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
    
    Call LoadProduct
    
    If p_status_prod_1 = True Then
        FillListview p_eng_product_1, lvList
        lblProd(0).Caption = p_prod_name_1
        Frame1.Enabled = True
    Else
        Frame1.Enabled = False
    End If
   
    If p_status_prod_2 = True Then
        FillListview p_eng_product_2, LvList2
        lblProd(1).Caption = p_prod_name_2
        Frame2.Enabled = True
    Else
        Frame2.Enabled = False
    End If
    
    itmrClose.InactiveInterval = 1000 * 10
    itmrClose.Enabled = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmInputSPB2 = Nothing
    Set RS_NG = Nothing
End Sub

Private Sub FillListview(eng_prod As String, listview As listview)
    Dim i As Integer
    Dim Rs As New Recordset
    Dim sSQL As String
    i = 1
    Rs.CursorLocation = adUseClient
    
    sSQL = "SELECT @no:=@no+1 AS nomor,a.id, a.plant_mark,a.date,a.period_shift,b.number,b.name as machine_name, "
    sSQL = sSQL & " b.tonnage,c.name as customer_name, d.internal_part_id,d.name as product_name,"
    sSQL = sSQL & " a.qty FROM sip_production.prod_spb2s_logs a JOIN (SELECT @no:=0) as no"
    sSQL = sSQL & " INNER JOIN sip_234.prod_machines as b  on a.prod_machine_id = b.id"
    sSQL = sSQL & " INNER JOIN sip_234.mkt_customers as c  on a.mkt_customer_id = c.id"
    sSQL = sSQL & " INNER JOIN sip_234.eng_products as d  on a.eng_product_id = d.id"
    sSQL = sSQL & " WHERE "
    sSQL = sSQL & " a.plant_mark = '" & p_plant_mark & "' "
    sSQL = sSQL & " and a.prod_machine_id = '" & p_prod_machine_id & "'"
    sSQL = sSQL & " and a.mkt_customer_id = '" & p_mkt_customer_id & "'"
    sSQL = sSQL & " and a.eng_product_id = '" & eng_prod & "'"
    sSQL = sSQL & " and date(a.created_at) = '" & Format(Date, "yyyy-mm-dd") & "'"
    sSQL = sSQL & " ORDER BY nomor DESC "
    
    Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
    
With listview
    
    .GridLines = True
    .View = lvwReport
    .SortOrder = lvwDescending
    
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "NO."
    .ColumnHeaders.Add , , "ID"
    .ColumnHeaders.Add , , "TANGGAL"
    .ColumnHeaders.Add , , "INTERNAL PART"
    .ColumnHeaders.Add , , "QTY"
    .ListItems.Clear
    Do While Not Rs.EOF
    Set srcItem = .ListItems.Add(, , Rs.Fields("nomor"), 1, 1)
        srcItem.SubItems(1) = Rs.Fields("id")
        srcItem.SubItems(2) = Format(Rs.Fields("date"), "yyyy-mm-dd hh:mm:ss")
        srcItem.SubItems(3) = Rs.Fields("internal_part_id")
        srcItem.SubItems(4) = Rs.Fields("qty")

        Rs.MoveNext
        
    Loop
        
End With
Call lvSizeColumns(listview)
End Sub


Private Sub itmrClose_UserInactive()
    Unload Me
End Sub

Private Sub LvList_DblClick()
    If MsgBox("Apakah anda yakin akan mengapus " & lvList.SelectedItem.SubItems(2) & " ?", vbQuestion + vbYesNo) = vbYes Then
        sSQL_Delete "DELETE FROM sip_production.prod_spb2s_logs WHERE id = " & lvList.SelectedItem.SubItems(1) & ""
        FillListview p_eng_product_1, lvList
    End If
End Sub


Private Sub LvList2_DblClick()
If MsgBox("Apakah anda yakin akan mengapus " & LvList2.SelectedItem.SubItems(3) & " ?", vbQuestion + vbYesNo) = vbYes Then
    sSQL_Delete "DELETE FROM sip_production.prod_spb2s_logs WHERE id = " & LvList2.SelectedItem.SubItems(1) & ""
    FillListview p_eng_product_2, LvList2
End If
End Sub
