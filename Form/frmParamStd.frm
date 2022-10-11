VERSION 5.00
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "b8Controls4.ocx"
Begin VB.Form frmParamStd 
   ClientHeight    =   12210
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   16695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12210
   ScaleWidth      =   16695
   Begin b8Controls4.b8TitleBar b8TitleBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   349
      Top             =   45
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   661
      Caption         =   "Parameter Standard"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      BackColor       =   8421504
   End
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   0
      TabIndex        =   334
      Top             =   495
      Width           =   11400
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
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
         Index           =   6
         Left            =   9315
         TabIndex        =   347
         Top             =   540
         Width           =   1905
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
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
         Index           =   5
         Left            =   6615
         TabIndex        =   345
         Top             =   540
         Width           =   1410
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
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
         Index           =   4
         Left            =   6615
         TabIndex        =   343
         Top             =   225
         Width           =   4605
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
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
         Index           =   3
         Left            =   3240
         TabIndex        =   341
         Top             =   540
         Width           =   1860
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
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
         Left            =   990
         TabIndex        =   337
         Top             =   225
         Width           =   1230
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
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
         Left            =   990
         TabIndex        =   336
         Top             =   540
         Width           =   1230
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
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
         Index           =   2
         Left            =   3240
         TabIndex        =   335
         Top             =   225
         Width           =   1860
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   196
         Left            =   8010
         TabIndex        =   348
         Top             =   540
         Width           =   1320
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rev."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   195
         Left            =   5130
         TabIndex        =   346
         Top             =   540
         Width           =   1500
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   194
         Left            =   5130
         TabIndex        =   344
         Top             =   225
         Width           =   1500
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tgl. Revisi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   193
         Left            =   2250
         TabIndex        =   342
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "M/c No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   192
         Left            =   225
         TabIndex        =   340
         Top             =   225
         Width           =   780
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Brand"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   191
         Left            =   225
         TabIndex        =   339
         Top             =   540
         Width           =   780
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Doc No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   190
         Left            =   2250
         TabIndex        =   338
         Top             =   225
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Index           =   8
      Left            =   11430
      TabIndex        =   284
      Top             =   6525
      Width           =   5145
      Begin VB.TextBox txtmaterial 
         Appearance      =   0  'Flat
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
         Index           =   15
         Left            =   3735
         TabIndex        =   331
         Top             =   2790
         Width           =   735
      End
      Begin VB.TextBox txtmaterial 
         Appearance      =   0  'Flat
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
         Index           =   14
         Left            =   3735
         TabIndex        =   328
         Top             =   2475
         Width           =   735
      End
      Begin VB.TextBox txtmaterial 
         Appearance      =   0  'Flat
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
         Index           =   13
         Left            =   3735
         TabIndex        =   325
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtmaterial 
         Appearance      =   0  'Flat
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
         Index           =   12
         Left            =   3735
         TabIndex        =   322
         Top             =   1845
         Width           =   735
      End
      Begin VB.TextBox txtmaterial 
         Appearance      =   0  'Flat
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
         Index           =   11
         Left            =   3735
         TabIndex        =   319
         Top             =   1530
         Width           =   735
      End
      Begin VB.TextBox txtmaterial 
         Appearance      =   0  'Flat
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
         Index           =   10
         Left            =   3735
         TabIndex        =   316
         Top             =   1215
         Width           =   735
      End
      Begin VB.TextBox txtmaterial 
         Appearance      =   0  'Flat
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
         Index           =   9
         Left            =   3735
         TabIndex        =   313
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtmaterial 
         Appearance      =   0  'Flat
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
         Index           =   8
         Left            =   3735
         TabIndex        =   310
         Top             =   585
         Width           =   735
      End
      Begin VB.TextBox txtmaterial 
         Appearance      =   0  'Flat
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
         Index           =   7
         Left            =   1170
         TabIndex        =   292
         Top             =   2790
         Width           =   870
      End
      Begin VB.TextBox txtmaterial 
         Appearance      =   0  'Flat
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
         Index           =   6
         Left            =   1170
         TabIndex        =   291
         Top             =   2475
         Width           =   870
      End
      Begin VB.TextBox txtmaterial 
         Appearance      =   0  'Flat
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
         Index           =   5
         Left            =   1170
         TabIndex        =   290
         Top             =   2160
         Width           =   870
      End
      Begin VB.TextBox txtmaterial 
         Appearance      =   0  'Flat
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
         Index           =   4
         Left            =   1170
         TabIndex        =   289
         Top             =   1845
         Width           =   870
      End
      Begin VB.TextBox txtmaterial 
         Appearance      =   0  'Flat
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
         Index           =   3
         Left            =   1170
         TabIndex        =   288
         Top             =   1530
         Width           =   870
      End
      Begin VB.TextBox txtmaterial 
         Appearance      =   0  'Flat
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
         Index           =   2
         Left            =   1170
         TabIndex        =   287
         Top             =   1215
         Width           =   870
      End
      Begin VB.TextBox txtmaterial 
         Appearance      =   0  'Flat
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
         Left            =   1170
         TabIndex        =   286
         Top             =   900
         Width           =   870
      End
      Begin VB.TextBox txtmaterial 
         Appearance      =   0  'Flat
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
         Left            =   1170
         TabIndex        =   285
         Top             =   585
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hot Runner 8"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   189
         Left            =   2565
         TabIndex        =   333
         Top             =   2790
         Width           =   1185
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   188
         Left            =   4455
         TabIndex        =   332
         Top             =   2790
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hot Runner 7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   158
         Left            =   2565
         TabIndex        =   330
         Top             =   2475
         Width           =   1185
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   157
         Left            =   4455
         TabIndex        =   329
         Top             =   2475
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hot Runner 6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   187
         Left            =   2565
         TabIndex        =   327
         Top             =   2160
         Width           =   1185
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   186
         Left            =   4455
         TabIndex        =   326
         Top             =   2160
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hot Runner 5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   185
         Left            =   2565
         TabIndex        =   324
         Top             =   1845
         Width           =   1185
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   184
         Left            =   4455
         TabIndex        =   323
         Top             =   1845
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hot Runner 4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   183
         Left            =   2565
         TabIndex        =   321
         Top             =   1530
         Width           =   1185
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   182
         Left            =   4455
         TabIndex        =   320
         Top             =   1530
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hot Runner 3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   181
         Left            =   2565
         TabIndex        =   318
         Top             =   1215
         Width           =   1185
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   180
         Left            =   4455
         TabIndex        =   317
         Top             =   1215
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hot Runner 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   179
         Left            =   2565
         TabIndex        =   315
         Top             =   900
         Width           =   1185
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   178
         Left            =   4455
         TabIndex        =   314
         Top             =   900
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hot Runner 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   168
         Left            =   2565
         TabIndex        =   312
         Top             =   585
         Width           =   1185
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   167
         Left            =   4455
         TabIndex        =   311
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Drying Time"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   177
         Left            =   180
         TabIndex        =   309
         Top             =   2790
         Width           =   1005
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hopper"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   176
         Left            =   180
         TabIndex        =   308
         Top             =   2475
         Width           =   1005
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "H4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   175
         Left            =   180
         TabIndex        =   307
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "H3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   174
         Left            =   180
         TabIndex        =   306
         Top             =   1845
         Width           =   1005
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "H2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   173
         Left            =   180
         TabIndex        =   305
         Top             =   1530
         Width           =   1005
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "H1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   172
         Left            =   180
         TabIndex        =   304
         Top             =   1215
         Width           =   1005
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   171
         Left            =   180
         TabIndex        =   303
         Top             =   900
         Width           =   1005
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MATERIAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   170
         Left            =   180
         TabIndex        =   302
         Top             =   270
         Width           =   4830
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HeN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   169
         Left            =   180
         TabIndex        =   301
         Top             =   585
         Width           =   1005
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   166
         Left            =   2025
         TabIndex        =   300
         Top             =   2790
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   165
         Left            =   2025
         TabIndex        =   299
         Top             =   2475
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   164
         Left            =   2025
         TabIndex        =   298
         Top             =   2160
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   163
         Left            =   2025
         TabIndex        =   297
         Top             =   1845
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   162
         Left            =   2025
         TabIndex        =   296
         Top             =   1530
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   161
         Left            =   2025
         TabIndex        =   295
         Top             =   1215
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   160
         Left            =   2025
         TabIndex        =   294
         Top             =   900
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   159
         Left            =   2025
         TabIndex        =   293
         Top             =   585
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4650
      Index           =   4
      Left            =   8010
      TabIndex        =   181
      Top             =   7110
      Width           =   3390
      Begin VB.TextBox txtSupporting 
         Appearance      =   0  'Flat
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
         Index           =   10
         Left            =   2385
         TabIndex        =   204
         Top             =   3735
         Width           =   870
      End
      Begin VB.TextBox txtSupporting 
         Appearance      =   0  'Flat
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
         Index           =   11
         Left            =   2385
         TabIndex        =   203
         Top             =   4050
         Width           =   870
      End
      Begin VB.TextBox txtSupporting 
         Appearance      =   0  'Flat
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
         Index           =   7
         Left            =   2385
         TabIndex        =   191
         Top             =   2790
         Width           =   870
      End
      Begin VB.TextBox txtSupporting 
         Appearance      =   0  'Flat
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
         Index           =   6
         Left            =   2385
         TabIndex        =   190
         Top             =   2475
         Width           =   870
      End
      Begin VB.TextBox txtSupporting 
         Appearance      =   0  'Flat
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
         Index           =   5
         Left            =   2385
         TabIndex        =   189
         Top             =   2160
         Width           =   870
      End
      Begin VB.TextBox txtSupporting 
         Appearance      =   0  'Flat
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
         Index           =   4
         Left            =   2385
         TabIndex        =   188
         Top             =   1845
         Width           =   870
      End
      Begin VB.TextBox txtSupporting 
         Appearance      =   0  'Flat
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
         Index           =   3
         Left            =   2385
         TabIndex        =   187
         Top             =   1530
         Width           =   870
      End
      Begin VB.TextBox txtSupporting 
         Appearance      =   0  'Flat
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
         Index           =   2
         Left            =   2385
         TabIndex        =   186
         Top             =   1215
         Width           =   870
      End
      Begin VB.TextBox txtSupporting 
         Appearance      =   0  'Flat
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
         Left            =   2385
         TabIndex        =   185
         Top             =   900
         Width           =   870
      End
      Begin VB.TextBox txtSupporting 
         Appearance      =   0  'Flat
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
         Left            =   2385
         TabIndex        =   184
         Top             =   585
         Width           =   870
      End
      Begin VB.TextBox txtSupporting 
         Appearance      =   0  'Flat
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
         Index           =   8
         Left            =   2385
         TabIndex        =   183
         Top             =   3105
         Width           =   870
      End
      Begin VB.TextBox txtSupporting 
         Appearance      =   0  'Flat
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
         Index           =   9
         Left            =   2385
         TabIndex        =   182
         Top             =   3420
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HOT RNR CTRL	"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   118
         Left            =   180
         TabIndex        =   212
         Top             =   3735
         Width           =   1095
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CHILLER"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   117
         Left            =   180
         TabIndex        =   211
         Top             =   3105
         Width           =   1095
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MTC	"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   116
         Left            =   180
         TabIndex        =   210
         Top             =   2475
         Width           =   1095
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Power Pack / Core Puller	"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   115
         Left            =   180
         TabIndex        =   209
         Top             =   1845
         Width           =   1095
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dehumidifier"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   111
         Left            =   180
         TabIndex        =   208
         Top             =   1215
         Width           =   1095
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hopper Dryer	"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   110
         Left            =   180
         TabIndex        =   207
         Top             =   585
         Width           =   1095
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "QTY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   109
         Left            =   1260
         TabIndex        =   206
         Top             =   4050
         Width           =   1140
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Model / Spec"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   108
         Left            =   1260
         TabIndex        =   205
         Top             =   3735
         Width           =   1140
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "QTY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   107
         Left            =   1260
         TabIndex        =   202
         Top             =   3420
         Width           =   1140
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Model / Spec"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   106
         Left            =   1260
         TabIndex        =   201
         Top             =   3105
         Width           =   1140
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "QTY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   105
         Left            =   1260
         TabIndex        =   200
         Top             =   2790
         Width           =   1140
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Model / Spec"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   104
         Left            =   1260
         TabIndex        =   199
         Top             =   2475
         Width           =   1140
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "QTY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   103
         Left            =   1260
         TabIndex        =   198
         Top             =   2160
         Width           =   1140
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Model / Spec"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   102
         Left            =   1260
         TabIndex        =   197
         Top             =   1845
         Width           =   1140
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "QTY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   101
         Left            =   1260
         TabIndex        =   196
         Top             =   1530
         Width           =   1140
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Model / Spec"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   100
         Left            =   1260
         TabIndex        =   195
         Top             =   1215
         Width           =   1140
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "QTY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   114
         Left            =   1260
         TabIndex        =   194
         Top             =   900
         Width           =   1140
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SUPPORTING"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   113
         Left            =   180
         TabIndex        =   193
         Top             =   270
         Width           =   3075
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Model / Spec"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   112
         Left            =   1260
         TabIndex        =   192
         Top             =   585
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4650
      Index           =   3
      Left            =   4860
      TabIndex        =   149
      Top             =   7110
      Width           =   3120
      Begin VB.TextBox txtmold 
         Appearance      =   0  'Flat
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
         Index           =   9
         Left            =   1485
         TabIndex        =   169
         Top             =   3420
         Width           =   915
      End
      Begin VB.TextBox txtmold 
         Appearance      =   0  'Flat
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
         Index           =   8
         Left            =   1485
         TabIndex        =   167
         Top             =   3105
         Width           =   915
      End
      Begin VB.TextBox txtmold 
         Appearance      =   0  'Flat
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
         Left            =   1485
         TabIndex        =   157
         Top             =   585
         Width           =   915
      End
      Begin VB.TextBox txtmold 
         Appearance      =   0  'Flat
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
         Left            =   1485
         TabIndex        =   156
         Top             =   900
         Width           =   915
      End
      Begin VB.TextBox txtmold 
         Appearance      =   0  'Flat
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
         Index           =   2
         Left            =   1485
         TabIndex        =   155
         Top             =   1215
         Width           =   915
      End
      Begin VB.TextBox txtmold 
         Appearance      =   0  'Flat
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
         Index           =   3
         Left            =   1485
         TabIndex        =   154
         Top             =   1530
         Width           =   915
      End
      Begin VB.TextBox txtmold 
         Appearance      =   0  'Flat
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
         Index           =   4
         Left            =   1485
         TabIndex        =   153
         Top             =   1845
         Width           =   915
      End
      Begin VB.TextBox txtmold 
         Appearance      =   0  'Flat
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
         Index           =   5
         Left            =   1485
         TabIndex        =   152
         Top             =   2160
         Width           =   915
      End
      Begin VB.TextBox txtmold 
         Appearance      =   0  'Flat
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
         Index           =   6
         Left            =   1485
         TabIndex        =   151
         Top             =   2475
         Width           =   915
      End
      Begin VB.TextBox txtmold 
         Appearance      =   0  'Flat
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
         Index           =   7
         Left            =   1485
         TabIndex        =   150
         Top             =   2790
         Width           =   915
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   99
         Left            =   2385
         TabIndex        =   180
         Top             =   3420
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   98
         Left            =   2385
         TabIndex        =   179
         Top             =   3105
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Plate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   97
         Left            =   2385
         TabIndex        =   178
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   96
         Left            =   2385
         TabIndex        =   177
         Top             =   900
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   95
         Left            =   2385
         TabIndex        =   176
         Top             =   1215
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   94
         Left            =   2385
         TabIndex        =   175
         Top             =   1530
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   93
         Left            =   2385
         TabIndex        =   174
         Top             =   1845
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   92
         Left            =   2385
         TabIndex        =   173
         Top             =   2160
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   91
         Left            =   2385
         TabIndex        =   172
         Top             =   2475
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   90
         Left            =   2385
         TabIndex        =   171
         Top             =   2790
         Width           =   555
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hold Sprue D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   89
         Left            =   180
         TabIndex        =   170
         Top             =   3420
         Width           =   1320
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "D Sprue"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   88
         Left            =   180
         TabIndex        =   168
         Top             =   3105
         Width           =   1320
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mold Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   87
         Left            =   180
         TabIndex        =   166
         Top             =   585
         Width           =   1320
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MOLD DATA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   86
         Left            =   180
         TabIndex        =   165
         Top             =   270
         Width           =   2760
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Slider"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   85
         Left            =   180
         TabIndex        =   164
         Top             =   900
         Width           =   1320
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Core Puller"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   84
         Left            =   180
         TabIndex        =   163
         Top             =   1215
         Width           =   1320
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Air Blow"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   83
         Left            =   180
         TabIndex        =   162
         Top             =   1530
         Width           =   1320
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hot Rnr Zone"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   82
         Left            =   180
         TabIndex        =   161
         Top             =   1845
         Width           =   1320
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ejector LS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   81
         Left            =   180
         TabIndex        =   160
         Top             =   2160
         Width           =   1320
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "V x H x T"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   80
         Left            =   180
         TabIndex        =   159
         Top             =   2475
         Width           =   1320
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R Sprue"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   79
         Left            =   180
         TabIndex        =   158
         Top             =   2790
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2985
      Index           =   7
      Left            =   11430
      TabIndex        =   5
      Top             =   3510
      Width           =   5145
      Begin VB.TextBox txtInjection 
         Appearance      =   0  'Flat
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
         Index           =   12
         Left            =   4275
         TabIndex        =   282
         Top             =   2205
         Width           =   735
      End
      Begin VB.TextBox txtInjection 
         Appearance      =   0  'Flat
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
         Index           =   11
         Left            =   4275
         TabIndex        =   280
         Top             =   1890
         Width           =   735
      End
      Begin VB.TextBox txtInjection 
         Appearance      =   0  'Flat
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
         Index           =   10
         Left            =   4275
         TabIndex        =   278
         Top             =   1575
         Width           =   735
      End
      Begin VB.TextBox txtInjection 
         Appearance      =   0  'Flat
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
         Index           =   9
         Left            =   4275
         TabIndex        =   276
         Top             =   1260
         Width           =   735
      End
      Begin VB.TextBox txtInjection 
         Appearance      =   0  'Flat
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
         Index           =   8
         Left            =   4275
         TabIndex        =   274
         Top             =   945
         Width           =   735
      End
      Begin VB.TextBox txtInjection 
         Appearance      =   0  'Flat
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
         Index           =   7
         Left            =   4275
         TabIndex        =   272
         Top             =   630
         Width           =   735
      End
      Begin VB.TextBox txtInjection 
         Appearance      =   0  'Flat
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
         Index           =   6
         Left            =   1710
         TabIndex        =   270
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtInjection 
         Appearance      =   0  'Flat
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
         Index           =   5
         Left            =   1710
         TabIndex        =   268
         Top             =   2205
         Width           =   735
      End
      Begin VB.TextBox txtInjection 
         Appearance      =   0  'Flat
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
         Index           =   4
         Left            =   1710
         TabIndex        =   266
         Top             =   1890
         Width           =   735
      End
      Begin VB.TextBox txtInjection 
         Appearance      =   0  'Flat
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
         Index           =   3
         Left            =   1710
         TabIndex        =   264
         Top             =   1575
         Width           =   735
      End
      Begin VB.TextBox txtInjection 
         Appearance      =   0  'Flat
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
         Index           =   2
         Left            =   1710
         TabIndex        =   262
         Top             =   1260
         Width           =   735
      End
      Begin VB.TextBox txtInjection 
         Appearance      =   0  'Flat
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
         Left            =   1710
         TabIndex        =   260
         Top             =   945
         Width           =   735
      End
      Begin VB.TextBox txtInjection 
         Appearance      =   0  'Flat
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
         Left            =   1710
         TabIndex        =   258
         Top             =   630
         Width           =   735
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "	T Slot / Biasa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   156
         Left            =   2430
         TabIndex        =   283
         Top             =   2205
         Width           =   1860
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Max Mould Thickness"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   155
         Left            =   2430
         TabIndex        =   281
         Top             =   1890
         Width           =   1860
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Min Mould Thickness"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   154
         Left            =   2430
         TabIndex        =   279
         Top             =   1575
         Width           =   1860
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Core Pack"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   153
         Left            =   2430
         TabIndex        =   277
         Top             =   1260
         Width           =   1860
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nozzle Diameter (mm)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   152
         Left            =   2430
         TabIndex        =   275
         Top             =   945
         Width           =   1860
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nozzle Length (mm)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   151
         Left            =   2430
         TabIndex        =   273
         Top             =   630
         Width           =   1860
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tie Bar (mm)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   150
         Left            =   135
         TabIndex        =   271
         Top             =   2520
         Width           =   1590
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Max Shot (pp / gr)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   149
         Left            =   135
         TabIndex        =   269
         Top             =   2205
         Width           =   1590
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Robot Pick Up"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   148
         Left            =   135
         TabIndex        =   267
         Top             =   1890
         Width           =   1590
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "D Nozzle (mm)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   147
         Left            =   135
         TabIndex        =   265
         Top             =   1575
         Width           =   1590
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R Nozzle (mm)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   146
         Left            =   135
         TabIndex        =   263
         Top             =   1260
         Width           =   1590
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Toggle/Direct"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   145
         Left            =   135
         TabIndex        =   261
         Top             =   945
         Width           =   1590
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hydraulic/Elect"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   144
         Left            =   135
         TabIndex        =   259
         Top             =   630
         Width           =   1590
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Machine Injection"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   143
         Left            =   135
         TabIndex        =   257
         Top             =   315
         Width           =   4875
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1950
      Index           =   6
      Left            =   11430
      TabIndex        =   4
      Top             =   9810
      Width           =   5145
      Begin VB.TextBox txtCooling 
         Appearance      =   0  'Flat
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
         Index           =   5
         Left            =   3960
         TabIndex        =   255
         Top             =   1485
         Width           =   1050
      End
      Begin VB.TextBox txtCooling 
         Appearance      =   0  'Flat
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
         Index           =   4
         Left            =   3960
         TabIndex        =   253
         Top             =   1170
         Width           =   1050
      End
      Begin VB.TextBox txtCooling 
         Appearance      =   0  'Flat
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
         Index           =   3
         Left            =   3960
         TabIndex        =   251
         Top             =   855
         Width           =   1050
      End
      Begin VB.TextBox txtCooling 
         Appearance      =   0  'Flat
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
         Index           =   2
         Left            =   1530
         TabIndex        =   249
         Top             =   1485
         Width           =   1005
      End
      Begin VB.TextBox txtCooling 
         Appearance      =   0  'Flat
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
         Left            =   1530
         TabIndex        =   247
         Top             =   1170
         Width           =   1005
      End
      Begin VB.TextBox txtCooling 
         Appearance      =   0  'Flat
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
         Left            =   1530
         TabIndex        =   242
         Top             =   855
         Width           =   1005
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "	COOLING TWR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   142
         Left            =   2520
         TabIndex        =   256
         Top             =   1485
         Width           =   1455
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CHILLER (C)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   141
         Left            =   2520
         TabIndex        =   254
         Top             =   1170
         Width           =   1455
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MTC (c)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   140
         Left            =   2520
         TabIndex        =   252
         Top             =   855
         Width           =   1455
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "COOLING TWR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   139
         Left            =   180
         TabIndex        =   250
         Top             =   1485
         Width           =   1365
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CHILLER (C)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   138
         Left            =   180
         TabIndex        =   248
         Top             =   1170
         Width           =   1365
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CORE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   137
         Left            =   2520
         TabIndex        =   246
         Top             =   540
         Width           =   2490
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CAVITY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   136
         Left            =   180
         TabIndex        =   245
         Top             =   540
         Width           =   2355
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MOLD COOLING"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   135
         Left            =   180
         TabIndex        =   244
         Top             =   180
         Width           =   4830
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MTC (c)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   134
         Left            =   180
         TabIndex        =   243
         Top             =   855
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2985
      Index           =   5
      Left            =   11430
      TabIndex        =   3
      Top             =   495
      Width           =   5145
      Begin VB.TextBox txtMonitoring 
         Appearance      =   0  'Flat
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
         Index           =   13
         Left            =   3960
         TabIndex        =   240
         Top             =   2520
         Width           =   1050
      End
      Begin VB.TextBox txtMonitoring 
         Appearance      =   0  'Flat
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
         Index           =   12
         Left            =   3960
         TabIndex        =   238
         Top             =   2205
         Width           =   1050
      End
      Begin VB.TextBox txtMonitoring 
         Appearance      =   0  'Flat
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
         Index           =   11
         Left            =   3960
         TabIndex        =   236
         Top             =   1890
         Width           =   1050
      End
      Begin VB.TextBox txtMonitoring 
         Appearance      =   0  'Flat
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
         Index           =   10
         Left            =   3960
         TabIndex        =   234
         Top             =   1575
         Width           =   1050
      End
      Begin VB.TextBox txtMonitoring 
         Appearance      =   0  'Flat
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
         Index           =   9
         Left            =   3960
         TabIndex        =   232
         Top             =   1260
         Width           =   1050
      End
      Begin VB.TextBox txtMonitoring 
         Appearance      =   0  'Flat
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
         Index           =   8
         Left            =   3960
         TabIndex        =   230
         Top             =   945
         Width           =   1050
      End
      Begin VB.TextBox txtMonitoring 
         Appearance      =   0  'Flat
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
         Index           =   7
         Left            =   3960
         TabIndex        =   228
         Top             =   630
         Width           =   1050
      End
      Begin VB.TextBox txtMonitoring 
         Appearance      =   0  'Flat
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
         Index           =   6
         Left            =   1395
         TabIndex        =   226
         Top             =   2520
         Width           =   870
      End
      Begin VB.TextBox txtMonitoring 
         Appearance      =   0  'Flat
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
         Index           =   5
         Left            =   1395
         TabIndex        =   224
         Top             =   2205
         Width           =   870
      End
      Begin VB.TextBox txtMonitoring 
         Appearance      =   0  'Flat
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
         Index           =   4
         Left            =   1395
         TabIndex        =   222
         Top             =   1890
         Width           =   870
      End
      Begin VB.TextBox txtMonitoring 
         Appearance      =   0  'Flat
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
         Index           =   3
         Left            =   1395
         TabIndex        =   220
         Top             =   1575
         Width           =   870
      End
      Begin VB.TextBox txtMonitoring 
         Appearance      =   0  'Flat
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
         Index           =   2
         Left            =   1395
         TabIndex        =   218
         Top             =   1260
         Width           =   870
      End
      Begin VB.TextBox txtMonitoring 
         Appearance      =   0  'Flat
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
         Left            =   1395
         TabIndex        =   216
         Top             =   945
         Width           =   870
      End
      Begin VB.TextBox txtMonitoring 
         Appearance      =   0  'Flat
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
         Left            =   1395
         TabIndex        =   213
         Top             =   630
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Act. Clm PCH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   133
         Left            =   2250
         TabIndex        =   241
         Top             =   2520
         Width           =   1725
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mld Pos Inject (mm)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   132
         Left            =   2250
         TabIndex        =   239
         Top             =   2205
         Width           =   1725
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mld Pos PCH (mm)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   131
         Left            =   2250
         TabIndex        =   237
         Top             =   1890
         Width           =   1725
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inj. Start Post (mm)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   130
         Left            =   2250
         TabIndex        =   235
         Top             =   1575
         Width           =   1725
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FPC Pres"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   129
         Left            =   2250
         TabIndex        =   233
         Top             =   1260
         Width           =   1725
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ACT-CUSH (mm)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   128
         Left            =   2250
         TabIndex        =   231
         Top             =   945
         Width           =   1725
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MIN-CUSH (mm)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   127
         Left            =   2250
         TabIndex        =   229
         Top             =   630
         Width           =   1725
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Act. Inj Pres"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   126
         Left            =   180
         TabIndex        =   227
         Top             =   2520
         Width           =   1230
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Act. Inj BP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   125
         Left            =   180
         TabIndex        =   225
         Top             =   2205
         Width           =   1230
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SCREW (rpm)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   124
         Left            =   180
         TabIndex        =   223
         Top             =   1890
         Width           =   1230
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CYCLE (s)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   123
         Left            =   180
         TabIndex        =   221
         Top             =   1575
         Width           =   1230
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TAKE (S)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   122
         Left            =   180
         TabIndex        =   219
         Top             =   1260
         Width           =   1230
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CHARGE (s)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   121
         Left            =   180
         TabIndex        =   217
         Top             =   945
         Width           =   1230
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MONITORING"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   120
         Left            =   180
         TabIndex        =   215
         Top             =   315
         Width           =   4830
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FILL (s)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   119
         Left            =   180
         TabIndex        =   214
         Top             =   630
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4650
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   7110
      Width           =   4830
      Begin VB.TextBox txtproduct 
         Appearance      =   0  'Flat
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
         Index           =   7
         Left            =   1665
         TabIndex        =   147
         Top             =   2790
         Width           =   2985
      End
      Begin VB.TextBox txtproduct 
         Appearance      =   0  'Flat
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
         Index           =   6
         Left            =   1665
         TabIndex        =   146
         Top             =   2475
         Width           =   2985
      End
      Begin VB.TextBox txtproduct 
         Appearance      =   0  'Flat
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
         Index           =   5
         Left            =   1665
         TabIndex        =   144
         Top             =   2160
         Width           =   2985
      End
      Begin VB.TextBox txtproduct 
         Appearance      =   0  'Flat
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
         Index           =   4
         Left            =   1665
         TabIndex        =   142
         Top             =   1845
         Width           =   2985
      End
      Begin VB.TextBox txtproduct 
         Appearance      =   0  'Flat
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
         Index           =   3
         Left            =   1665
         TabIndex        =   140
         Top             =   1530
         Width           =   2985
      End
      Begin VB.TextBox txtproduct 
         Appearance      =   0  'Flat
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
         Index           =   2
         Left            =   1665
         TabIndex        =   138
         Top             =   1215
         Width           =   2985
      End
      Begin VB.TextBox txtproduct 
         Appearance      =   0  'Flat
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
         Left            =   1665
         TabIndex        =   136
         Top             =   900
         Width           =   2985
      End
      Begin VB.TextBox txtproduct 
         Appearance      =   0  'Flat
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
         Left            =   1665
         TabIndex        =   133
         Top             =   585
         Width           =   2985
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Shot Target/Jam"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   78
         Left            =   180
         TabIndex        =   148
         Top             =   2790
         Width           =   1500
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cavity No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   77
         Left            =   180
         TabIndex        =   145
         Top             =   2475
         Width           =   1500
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Runner (gr)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   76
         Left            =   180
         TabIndex        =   143
         Top             =   2160
         Width           =   1500
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Part (gr)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   75
         Left            =   180
         TabIndex        =   141
         Top             =   1845
         Width           =   1500
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   74
         Left            =   180
         TabIndex        =   139
         Top             =   1530
         Width           =   1500
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Material"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   73
         Left            =   180
         TabIndex        =   137
         Top             =   1215
         Width           =   1500
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Part No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   72
         Left            =   180
         TabIndex        =   135
         Top             =   900
         Width           =   1500
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRODUCT DATA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   71
         Left            =   180
         TabIndex        =   134
         Top             =   270
         Width           =   4470
      End
      Begin VB.Label lblInjectrol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Part Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   70
         Left            =   180
         TabIndex        =   132
         Top             =   585
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3300
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   3780
      Width           =   11400
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   31
         Left            =   9540
         TabIndex        =   130
         Text            =   "0.0"
         Top             =   2745
         Width           =   1320
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   30
         Left            =   8235
         TabIndex        =   128
         Text            =   "0.0"
         Top             =   2745
         Width           =   1320
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   29
         Left            =   6930
         TabIndex        =   126
         Text            =   "0.0"
         Top             =   2745
         Width           =   1320
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   28
         Left            =   5310
         TabIndex        =   123
         Text            =   "0.0"
         Top             =   2835
         Width           =   1455
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   27
         Left            =   5310
         TabIndex        =   121
         Text            =   "0.0"
         Top             =   2250
         Width           =   1455
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   26
         Left            =   2745
         TabIndex        =   118
         Text            =   "0.0"
         Top             =   2835
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   25
         Left            =   1035
         TabIndex        =   116
         Text            =   "0.0"
         Top             =   2835
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   24
         Left            =   180
         TabIndex        =   114
         Text            =   "0.0"
         Top             =   2835
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   23
         Left            =   2745
         TabIndex        =   112
         Text            =   "0.0"
         Top             =   2250
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   22
         Left            =   1035
         TabIndex        =   110
         Text            =   "0.0"
         Top             =   2250
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   21
         Left            =   180
         TabIndex        =   108
         Text            =   "0.0"
         Top             =   2250
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   20
         Left            =   9315
         TabIndex        =   103
         Text            =   "0.0"
         Top             =   1575
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   19
         Left            =   10170
         TabIndex        =   101
         Text            =   "0.0"
         Top             =   1575
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   18
         Left            =   7605
         TabIndex        =   99
         Text            =   "0.0"
         Top             =   1575
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   17
         Left            =   6750
         TabIndex        =   97
         Text            =   "0.0"
         Top             =   1575
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   16
         Left            =   5310
         TabIndex        =   95
         Text            =   "0.0"
         Top             =   1575
         Width           =   1455
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   15
         Left            =   4455
         TabIndex        =   93
         Text            =   "0.0"
         Top             =   1575
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   14
         Left            =   3600
         TabIndex        =   91
         Text            =   "0.0"
         Top             =   1575
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   13
         Left            =   2745
         TabIndex        =   89
         Text            =   "0.0"
         Top             =   1575
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   12
         Left            =   1890
         TabIndex        =   87
         Text            =   "0.0"
         Top             =   1575
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   11
         Left            =   1035
         TabIndex        =   85
         Text            =   "0.0"
         Top             =   1575
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   10
         Left            =   180
         TabIndex        =   83
         Text            =   "0.0"
         Top             =   1575
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   9
         Left            =   9315
         TabIndex        =   81
         Text            =   "0.0"
         Top             =   900
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   8
         Left            =   8460
         TabIndex        =   79
         Text            =   "0.0"
         Top             =   900
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   7
         Left            =   7605
         TabIndex        =   77
         Text            =   "0.0"
         Top             =   900
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   6
         Left            =   6750
         TabIndex        =   75
         Text            =   "0.0"
         Top             =   900
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   5
         Left            =   5310
         TabIndex        =   73
         Text            =   "0.0"
         Top             =   900
         Width           =   1455
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   4
         Left            =   3600
         TabIndex        =   71
         Text            =   "0.0"
         Top             =   900
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   3
         Left            =   2745
         TabIndex        =   69
         Text            =   "0.0"
         Top             =   900
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   2
         Left            =   1890
         TabIndex        =   67
         Text            =   "0.0"
         Top             =   900
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   1035
         TabIndex        =   65
         Text            =   "0.0"
         Top             =   900
         Width           =   870
      End
      Begin VB.TextBox txtClamprol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   180
         TabIndex        =   63
         Text            =   "0.0"
         Top             =   900
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TIMER"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   69
         Left            =   6930
         TabIndex        =   131
         Top             =   2205
         Width           =   3930
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   1185
         Left            =   6750
         Top             =   1980
         Width           =   4290
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inject Time"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   68
         Left            =   9540
         TabIndex        =   129
         Top             =   2475
         Width           =   1320
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cooling Time"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   67
         Left            =   8235
         TabIndex        =   127
         Top             =   2475
         Width           =   1320
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Interval"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   66
         Left            =   6930
         TabIndex        =   125
         Top             =   2475
         Width           =   1320
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1185
         Index           =   65
         Left            =   3600
         TabIndex        =   124
         Top             =   1980
         Width           =   1725
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ejc.Mode Off"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   64
         Left            =   5310
         TabIndex        =   122
         Top             =   2565
         Width           =   1455
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ejc.Ctn/Hld"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   63
         Left            =   5310
         TabIndex        =   120
         Top             =   1980
         Width           =   1455
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1185
         Index           =   62
         Left            =   1890
         TabIndex        =   119
         Top             =   1980
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LS 32"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   61
         Left            =   2745
         TabIndex        =   117
         Top             =   2565
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LS 31"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   60
         Left            =   1035
         TabIndex        =   115
         Top             =   2565
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LS 31A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   59
         Left            =   180
         TabIndex        =   113
         Top             =   2565
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   58
         Left            =   2745
         TabIndex        =   111
         Top             =   1980
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VE 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   57
         Left            =   1035
         TabIndex        =   109
         Top             =   1980
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VE 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   56
         Left            =   180
         TabIndex        =   107
         Top             =   1980
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   55
         Left            =   8460
         TabIndex        =   106
         Top             =   1305
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   54
         Left            =   10170
         TabIndex        =   105
         Top             =   630
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   600
         Index           =   53
         Left            =   4455
         TabIndex        =   104
         Top             =   630
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PCL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   52
         Left            =   9315
         TabIndex        =   102
         Top             =   1305
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PCH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   51
         Left            =   10170
         TabIndex        =   100
         Top             =   1305
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LS 2A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   50
         Left            =   7605
         TabIndex        =   98
         Top             =   1305
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LS 2E"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   49
         Left            =   6750
         TabIndex        =   96
         Top             =   1305
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LS 2D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   48
         Left            =   5310
         TabIndex        =   94
         Top             =   1305
         Width           =   1455
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LS 3M"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   47
         Left            =   4455
         TabIndex        =   92
         Top             =   1305
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LS 3A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   46
         Left            =   3600
         TabIndex        =   90
         Top             =   1305
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LS 3D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   45
         Left            =   2745
         TabIndex        =   88
         Top             =   1305
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LS 3E"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   44
         Left            =   1890
         TabIndex        =   86
         Top             =   1305
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LS 3B"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   43
         Left            =   1035
         TabIndex        =   84
         Top             =   1305
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LS 3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   42
         Left            =   180
         TabIndex        =   82
         Top             =   1305
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LS-2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   41
         Left            =   9315
         TabIndex        =   80
         Top             =   630
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VC S"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   40
         Left            =   8460
         TabIndex        =   78
         Top             =   630
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VC 3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   39
         Left            =   7605
         TabIndex        =   76
         Top             =   630
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VC 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   38
         Left            =   6750
         TabIndex        =   74
         Top             =   630
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VC 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   37
         Left            =   5310
         TabIndex        =   72
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VOL S1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   36
         Left            =   3600
         TabIndex        =   70
         Top             =   630
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VOL 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   35
         Left            =   2745
         TabIndex        =   68
         Top             =   630
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VOL 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   34
         Left            =   1890
         TabIndex        =   66
         Top             =   630
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VOL 3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   33
         Left            =   1035
         TabIndex        =   64
         Top             =   630
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VO S2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   32
         Left            =   180
         TabIndex        =   62
         Top             =   630
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CLAMPROL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   31
         Left            =   180
         TabIndex        =   61
         Top             =   360
         Width           =   10860
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2220
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   1530
      Width           =   11400
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   23
         Left            =   1035
         TabIndex        =   57
         Text            =   "0.0"
         Top             =   1755
         Width           =   870
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   22
         Left            =   10305
         TabIndex        =   46
         Text            =   "0.0"
         Top             =   1755
         Width           =   960
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   21
         Left            =   9360
         TabIndex        =   45
         Text            =   "0.0"
         Top             =   1755
         Width           =   960
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   20
         Left            =   8730
         TabIndex        =   44
         Text            =   "0.0"
         Top             =   1755
         Width           =   645
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   19
         Left            =   7875
         TabIndex        =   43
         Text            =   "0.0"
         Top             =   1755
         Width           =   870
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   18
         Left            =   7020
         TabIndex        =   42
         Text            =   "0.0"
         Top             =   1755
         Width           =   870
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   17
         Left            =   6165
         TabIndex        =   41
         Text            =   "0.0"
         Top             =   1755
         Width           =   870
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   16
         Left            =   5310
         TabIndex        =   40
         Text            =   "0.0"
         Top             =   1755
         Width           =   870
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   15
         Left            =   4455
         TabIndex        =   39
         Text            =   "0.0"
         Top             =   1755
         Width           =   870
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   14
         Left            =   2745
         TabIndex        =   38
         Text            =   "0.0"
         Top             =   1755
         Width           =   870
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   13
         Left            =   1890
         TabIndex        =   37
         Text            =   "0.0"
         Top             =   1755
         Width           =   870
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   12
         Left            =   8730
         TabIndex        =   33
         Text            =   "0.0"
         Top             =   1125
         Width           =   645
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   11
         Left            =   10305
         TabIndex        =   17
         Text            =   "0.0"
         Top             =   1125
         Width           =   960
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   10
         Left            =   9360
         TabIndex        =   16
         Text            =   "0.0"
         Top             =   1125
         Width           =   960
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   9
         Left            =   7875
         TabIndex        =   15
         Text            =   "0.0"
         Top             =   1125
         Width           =   870
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   8
         Left            =   7020
         TabIndex        =   14
         Text            =   "0.0"
         Top             =   1125
         Width           =   870
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   7
         Left            =   6165
         TabIndex        =   13
         Text            =   "0.0"
         Top             =   1125
         Width           =   870
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   6
         Left            =   5310
         TabIndex        =   12
         Text            =   "0.0"
         Top             =   1125
         Width           =   870
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   5
         Left            =   4455
         TabIndex        =   11
         Text            =   "0.0"
         Top             =   1125
         Width           =   870
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   4
         Left            =   3600
         TabIndex        =   10
         Text            =   "0.0"
         Top             =   1125
         Width           =   870
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   330
         Index           =   3
         Left            =   2745
         TabIndex        =   9
         Text            =   "0.0"
         Top             =   1125
         Width           =   870
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Index           =   2
         Left            =   1890
         TabIndex        =   8
         Text            =   "0.0"
         Top             =   1125
         Width           =   870
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   1035
         TabIndex        =   7
         Text            =   "0.0"
         Top             =   1125
         Width           =   870
      End
      Begin VB.TextBox txtInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   180
         TabIndex        =   6
         Text            =   "0.0"
         Top             =   1125
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   30
         Left            =   3600
         TabIndex        =   60
         Top             =   1530
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   29
         Left            =   180
         TabIndex        =   59
         Top             =   1530
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TRH-3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   28
         Left            =   1035
         TabIndex        =   58
         Top             =   1530
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Back Press"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   27
         Left            =   10305
         TabIndex        =   56
         Top             =   1530
         Width           =   960
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LS-10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   26
         Left            =   9360
         TabIndex        =   55
         Top             =   1530
         Width           =   960
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LS-4D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   25
         Left            =   5310
         TabIndex        =   54
         Top             =   1530
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LS-4C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   24
         Left            =   6165
         TabIndex        =   53
         Top             =   1530
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LS-4B"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   23
         Left            =   7020
         TabIndex        =   52
         Top             =   1530
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LS-4A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   22
         Left            =   7875
         TabIndex        =   51
         Top             =   1530
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LS-5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   21
         Left            =   8730
         TabIndex        =   50
         Top             =   1530
         Width           =   645
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LS-4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   20
         Left            =   4455
         TabIndex        =   49
         Top             =   1530
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TRH-1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   19
         Left            =   2745
         TabIndex        =   48
         Top             =   1530
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TRH-2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   18
         Left            =   1890
         TabIndex        =   47
         Top             =   1530
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INJECTROL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   17
         Left            =   180
         TabIndex        =   36
         Top             =   360
         Width           =   11085
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Charge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   16
         Left            =   9360
         TabIndex        =   35
         Top             =   630
         Width           =   1905
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PI-1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   15
         Left            =   8730
         TabIndex        =   34
         Top             =   630
         Width           =   645
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inj. Speed"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   14
         Left            =   4455
         TabIndex        =   32
         Top             =   630
         Width           =   4290
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hold. Spd"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   13
         Left            =   3600
         TabIndex        =   31
         Top             =   630
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hold. Press"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   12
         Left            =   180
         TabIndex        =   30
         Top             =   630
         Width           =   3435
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SRN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   11
         Left            =   10305
         TabIndex        =   29
         Top             =   900
         Width           =   960
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Low/High"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   10
         Left            =   9360
         TabIndex        =   28
         Top             =   900
         Width           =   960
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vi.1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   9
         Left            =   7875
         TabIndex        =   27
         Top             =   900
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vi.2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   7020
         TabIndex        =   26
         Top             =   900
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PH-1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   2745
         TabIndex        =   25
         Top             =   900
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VH 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   3600
         TabIndex        =   24
         Top             =   900
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vi.5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   4455
         TabIndex        =   23
         Top             =   900
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vi.4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   5310
         TabIndex        =   22
         Top             =   900
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vi.3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   6165
         TabIndex        =   21
         Top             =   900
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PH-2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   1890
         TabIndex        =   20
         Top             =   900
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PH-3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   1035
         TabIndex        =   19
         Top             =   900
         Width           =   870
      End
      Begin VB.Label lblInjectrol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PH-4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   18
         Top             =   900
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmParamStd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LoadData()
On Error GoTo ErrHandler

    Dim RS As New Recordset
    Dim sSQL As String

    RS.CursorLocation = adUseClient

    sSQL = "SELECT a.*,b.NAME AS customer_name, c.NUMBER AS machine_no,c.NAME AS machine_name,c.tonnage,"
    sSQL = sSQL & " d.internal_part_id,d.product_name,d.customer_part_number,d.material_name,d.color_name,"
    sSQL = sSQL & " d.cavity , d.weight_gr, d.weight_runner_gr, d.cycle_time_ia, d.target_shot"
    sSQL = sSQL & " FROM sip_production.eng_paramset_standards a"
    sSQL = sSQL & " LEFT JOIN sip_production.mkt_customers b ON a.mkt_customer_id = b.id"
    sSQL = sSQL & " LEFT JOIN sip_production.prod_machines c ON a.prod_machine_id = c.id"
    sSQL = sSQL & " LEFT JOIN (SELECT X.id,X.internal_part_id,X.NAME AS product_name,X.customer_part_number,X.cavity,X.weight_gr,X.weight_runner_gr,X.cycle_time_ia,"
                sSQL = sSQL & " ROUND((3600 / X.cycle_time_ia),0) AS target_shot,Y.NAME AS material_name, Z.NAME AS color_name FROM sip_production.eng_products X"
                sSQL = sSQL & " LEFT JOIN sip_production.eng_materials Y ON X.eng_material_for_label_id = Y.id"
                sSQL = sSQL & " LEFT JOIN sip_production.eng_colors Z ON X.eng_color_id = Z.id) d ON a.eng_product_id = d.id"
    sSQL = sSQL & " WHERE a.prod_machine_id = '" & p_prod_machine_id & "' AND d.internal_part_id = '" & p_int_part_1 & "'"

    RS.Open sSQL, CN, adOpenStatic, adLockOptimistic
    
    If RS.RecordCount > 0 Then
    
        txtInfo(0).Text = RS.Fields("machine_no")
        txtInfo(1).Text = RS.Fields("machine_name")
        txtInfo(2).Text = "ENG/PAR/" & Format(Now, "yy") & "/" & RS.Fields("number")
        txtInfo(3).Text = Format(RS.Fields("rev_date"), "yyyy-mm-dd")
        txtInfo(4).Text = RS.Fields("customer_name")
        txtInfo(5).Text = RS.Fields("rev")
        txtInfo(6).Text = Format(RS.Fields("date"), "yyyy-mm-dd")
        
        txtInjectrol(0).Text = RS.Fields("injectrol_ph_4")
        txtInjectrol(1).Text = RS.Fields("injectrol_ph_3")
        txtInjectrol(2).Text = RS.Fields("injectrol_ph_2")
        txtInjectrol(3).Text = RS.Fields("injectrol_ph_1")
        txtInjectrol(4).Text = RS.Fields("injectrol_vh_1")
        txtInjectrol(5).Text = RS.Fields("injectrol_vi_5")
        txtInjectrol(6).Text = RS.Fields("injectrol_vi_4")
        txtInjectrol(7).Text = RS.Fields("injectrol_vi_3")
        txtInjectrol(8).Text = RS.Fields("injectrol_vi_2")
        txtInjectrol(9).Text = RS.Fields("injectrol_vi_1")
        txtInjectrol(12).Text = RS.Fields("injectrol_pi_1")
        txtInjectrol(10).Text = IIf(IsNull(RS.Fields("machine_charge_type")), "", RS.Fields("machine_charge_type"))
        txtInjectrol(11).Text = RS.Fields("injectrol_srn")
        txtInjectrol(23).Text = RS.Fields("injectrol_trh_3")
        txtInjectrol(13).Text = RS.Fields("injectrol_trh_2")
        txtInjectrol(14).Text = RS.Fields("injectrol_trh_1")
        txtInjectrol(15).Text = RS.Fields("injectrol_ls_4")
        txtInjectrol(16).Text = RS.Fields("injectrol_ls_4d")
        txtInjectrol(17).Text = RS.Fields("injectrol_ls_4c")
        txtInjectrol(18).Text = RS.Fields("injectrol_ls_4b")
        txtInjectrol(19).Text = RS.Fields("injectrol_ls_5")
        txtInjectrol(21).Text = RS.Fields("injectrol_ls_10")
        txtInjectrol(22).Text = RS.Fields("injectrol_back_press")
        
        txtClamprol(29).Text = RS.Fields("injectrol_interval")
        txtClamprol(30).Text = RS.Fields("injectrol_cooling_time")
        txtClamprol(31).Text = RS.Fields("injectrol_inject_time")
        
        txtMonitoring(0).Text = RS.Fields("monitoring_fill")
        txtMonitoring(1).Text = RS.Fields("monitoring_charge")
        txtMonitoring(2).Text = RS.Fields("monitoring_takeout")
        txtMonitoring(3).Text = RS.Fields("monitoring_cycle")
        txtMonitoring(4).Text = RS.Fields("injectrol_screw_speed")
        
        txtMonitoring(4).Text = RS.Fields("monitoring_inj_peak")
        
        txtMonitoring(7).Text = RS.Fields("monitoring_min_cush")
        txtMonitoring(8).Text = RS.Fields("monitoring_act_cush")
        txtMonitoring(9).Text = RS.Fields("monitoring_fpc_press")
        txtMonitoring(10).Text = RS.Fields("monitoring_inj_start")
        
        txtMonitoring(13).Text = RS.Fields("clamprol_clamp_kn")

        txtClamprol(0).Text = RS.Fields("clamprol_vo_s2")
        txtClamprol(1).Text = RS.Fields("clamprol_vo_3")
        txtClamprol(2).Text = RS.Fields("clamprol_vo_2")
        txtClamprol(3).Text = RS.Fields("clamprol_vo_1")
        txtClamprol(4).Text = RS.Fields("clamprol_vo_s1")
        
        txtClamprol(5).Text = RS.Fields("clamprol_vc_1")
        txtClamprol(6).Text = RS.Fields("clamprol_vc_2")
        txtClamprol(7).Text = RS.Fields("clamprol_vc_3")
        txtClamprol(8).Text = RS.Fields("clamprol_vc_s")
        txtClamprol(9).Text = RS.Fields("clamprol_ls_2")
        
        txtClamprol(20).Text = RS.Fields("clamprol_pcl")
        txtClamprol(19).Text = RS.Fields("clamprol_pch")
        
        txtClamprol(10).Text = RS.Fields("clamprol_ls_3")
        txtClamprol(11).Text = RS.Fields("clamprol_ls_3b")
        txtClamprol(12).Text = RS.Fields("clamprol_ls_3e")
        txtClamprol(13).Text = RS.Fields("clamprol_ls_3d")
        txtClamprol(14).Text = RS.Fields("clamprol_ls_3a")
        txtClamprol(15).Text = RS.Fields("clamprol_ls_3m")
        txtClamprol(16).Text = RS.Fields("clamprol_ls_2d")
        txtClamprol(17).Text = RS.Fields("clamprol_ls_2e")
        txtClamprol(18).Text = RS.Fields("clamprol_ls_2a")
        
        txtClamprol(21).Text = RS.Fields("clamprol_ve_1")
        txtClamprol(22).Text = RS.Fields("clamprol_ve_2")
        txtClamprol(23).Text = RS.Fields("clamprol_vr")
        
        txtClamprol(28).Text = RS.Fields("clamprol_eject_mode")
        txtClamprol(27).Text = RS.Fields("clamprol_eject_count")
        
        txtClamprol(24).Text = RS.Fields("clamprol_ls_31a")
        txtClamprol(25).Text = RS.Fields("clamprol_ls_31")
        txtClamprol(26).Text = RS.Fields("clamprol_ls_32")
        
        txtproduct(0).Text = RS.Fields("product_name")
        txtproduct(1).Text = RS.Fields("customer_part_number")
        txtproduct(2).Text = RS.Fields("material_name")
        txtproduct(3).Text = IIf(IsNull(RS.Fields("color_name")), "", RS.Fields("color_name"))
        txtproduct(4).Text = RS.Fields("weight_gr")
        txtproduct(5).Text = RS.Fields("weight_runner_gr")
        txtproduct(6).Text = RS.Fields("cavity")
        txtproduct(7).Text = RS.Fields("target_shot")

        
        txtCooling(0).Text = RS.Fields("mold_cooling_cavity_mtc")
        txtCooling(1).Text = RS.Fields("mold_cooling_cavity_chiller")
        txtCooling(2).Text = RS.Fields("mold_cooling_cavity_cooling_twr")
        txtCooling(3).Text = RS.Fields("mold_cooling_core_mtc")
        txtCooling(4).Text = RS.Fields("mold_cooling_core_chiller")
        txtCooling(5).Text = RS.Fields("mold_cooling_core_cooling_twr")
        
        txtmold(0).Text = RS.Fields("mold_data_type")
        txtmold(1).Text = RS.Fields("mold_data_slider")
        txtmold(2).Text = RS.Fields("mold_data_core_puller")
        txtmold(3).Text = RS.Fields("mold_data_air_blow")
        txtmold(4).Text = RS.Fields("mold_data_hot_rnr_zone")
        txtmold(5).Text = RS.Fields("mold_data_ejector_ls")
        txtmold(6).Text = RS.Fields("mold_data_v") & " x " & RS.Fields("mold_data_h") & " x " & RS.Fields("mold_data_t")
        txtmold(7).Text = RS.Fields("mold_data_r_sprue")
        txtmold(8).Text = RS.Fields("mold_data_d_sprue")
        txtmold(9).Text = RS.Fields("mold_data_hole_sprue_d")
        
        txtmaterial(0).Text = RS.Fields("material_hen")
        txtmaterial(1).Text = RS.Fields("material_hn")
        txtmaterial(2).Text = RS.Fields("material_h1")
        txtmaterial(3).Text = RS.Fields("material_h2")
        txtmaterial(4).Text = RS.Fields("material_h3")
        txtmaterial(5).Text = RS.Fields("material_h4")
        
        txtmaterial(6).Text = RS.Fields("material_hopper_dehumi")
        txtmaterial(7).Text = RS.Fields("material_drying_time")
        
        txtmaterial(8).Text = RS.Fields("material_hot_runner_1")
        txtmaterial(9).Text = RS.Fields("material_hot_runner_2")
        txtmaterial(10).Text = RS.Fields("material_hot_runner_3")
        txtmaterial(11).Text = RS.Fields("material_hot_runner_4")
        txtmaterial(12).Text = RS.Fields("material_hot_runner_5")
        txtmaterial(13).Text = RS.Fields("material_hot_runner_6")
        txtmaterial(14).Text = RS.Fields("material_hot_runner_7")
        txtmaterial(15).Text = RS.Fields("material_hot_runner_8")
        
        txtSupporting(0).Text = RS.Fields("supporting_hopper_dryer_spec")
        txtSupporting(1).Text = RS.Fields("supporting_hopper_dryer_qty")
        txtSupporting(2).Text = RS.Fields("supporting_dehumidifier_spec")
        txtSupporting(3).Text = RS.Fields("supporting_dehumidifier_qty")
        txtSupporting(4).Text = RS.Fields("supporting_power_pack_spec")
        txtSupporting(5).Text = RS.Fields("supporting_power_pack_qty")
        txtSupporting(6).Text = RS.Fields("supporting_mtc_spec")
        txtSupporting(7).Text = RS.Fields("supporting_mtc_qty")
        txtSupporting(8).Text = RS.Fields("supporting_chiller_spec")
        txtSupporting(9).Text = RS.Fields("supporting_chiller_qty")
        txtSupporting(10).Text = RS.Fields("supporting_hot_rnr_ctrl_spec")
        txtSupporting(11).Text = RS.Fields("supporting_hot_rnr_ctrl_qty")
        
        txtInjection(0).Text = RS.Fields("machine_injection_type")
        txtInjection(1).Text = RS.Fields("machine_injection_clamping")
        txtInjection(2).Text = IIf(IsNull(RS.Fields("machine_injection_r_nozzle")), "", RS.Fields("machine_injection_r_nozzle"))
        txtInjection(3).Text = IIf(IsNull(RS.Fields("machine_injection_d_nozzle")), "", RS.Fields("machine_injection_d_nozzle"))
        txtInjection(4).Text = RS.Fields("machine_injection_robot_pick_up")
        txtInjection(5).Text = IIf(IsNull(RS.Fields("machine_injection_max_shot")), "", RS.Fields("machine_injection_max_shot"))
        txtInjection(6).Text = IIf(IsNull(RS.Fields("machine_injection_tie_bar_length")), "", RS.Fields("machine_injection_tie_bar_length"))
        txtInjection(7).Text = IIf(IsNull(RS.Fields("machine_injection_nozzle_length")), "", RS.Fields("machine_injection_nozzle_length"))
        txtInjection(8).Text = IIf(IsNull(RS.Fields("machine_injection_nozzle_diameter")), "", RS.Fields("machine_injection_nozzle_diameter"))
        txtInjection(9).Text = IIf(IsNull(RS.Fields("machine_injection_core_pack")), "", RS.Fields("machine_injection_core_pack"))
        txtInjection(10).Text = IIf(IsNull(RS.Fields("machine_injection_min_mould_thickness")), "", RS.Fields("machine_injection_min_mould_thickness"))
        txtInjection(11).Text = IIf(IsNull(RS.Fields("machine_injection_max_mould_thickness")), "", RS.Fields("machine_injection_max_mould_thickness"))
        txtInjection(12).Text = IIf(IsNull(RS.Fields("machine_injection_mold_clamping")), "", RS.Fields("machine_injection_mold_clamping"))
    
        
    
    End If
    
    Set RS = Nothing

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
    
   
End Sub

Private Sub Form_Load()
Call LoadData
End Sub

Private Sub Form_Resize()
On Error Resume Next

    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500
        b8TitleBar1.Width = Me.ScaleWidth
    End If
    
End Sub

