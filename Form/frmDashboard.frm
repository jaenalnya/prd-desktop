VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "B8CONT~2.OCX"
Begin VB.Form FrmDashboard 
   Caption         =   "Dashboard"
   ClientHeight    =   10500
   ClientLeft      =   3195
   ClientTop       =   1995
   ClientWidth     =   16455
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10500
   ScaleWidth      =   16455
   Begin VB.ComboBox CboProduct 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1710
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   855
      Width           =   8115
   End
   Begin VB.TextBox txtTarget_Yield_1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   10350
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "0"
      Top             =   5985
      Width           =   1950
   End
   Begin VB.TextBox txtPercTarget_1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   14445
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0"
      Top             =   5940
      Width           =   1770
   End
   Begin VB.TextBox txtTotalYield_1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   12510
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0"
      Top             =   5940
      Width           =   1770
   End
   Begin VB.TextBox txtgross_1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   10350
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0"
      Top             =   4950
      Width           =   1950
   End
   Begin VB.TextBox txtOk_1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   14445
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0"
      Top             =   4950
      Width           =   1770
   End
   Begin VB.TextBox txtNg_1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   12510
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0"
      Top             =   4950
      Width           =   1725
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   16200
      Top             =   360
   End
   Begin b8Controls4.b8TitleBar b8TitleBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16305
      _ExtentX        =   28760
      _ExtentY        =   661
      Caption         =   ""
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
   Begin MSComctlLib.ImageList i16x16 
      Left            =   135
      Top             =   8235
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
            Picture         =   "frmDashboard.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDashboard.frx":0A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDashboard.frx":1424
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDashboard.frx":17BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDashboard.frx":1B58
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDashboard.frx":1EF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDashboard.frx":228C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDashboard.frx":2C9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDashboard.frx":36B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDashboard.frx":40C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDashboard.frx":4AD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDashboard.frx":54E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDashboard.frx":5EF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDashboard.frx":690A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDashboard.frx":6EA6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   1770
      Left            =   180
      TabIndex        =   3
      Top             =   2565
      Width           =   16080
      _ExtentX        =   28363
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
   Begin MSComctlLib.ListView LvListIdle1 
      Height          =   3120
      Left            =   8550
      TabIndex        =   7
      Top             =   6885
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   5503
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   240
      Left            =   180
      TabIndex        =   11
      Top             =   10125
      Width           =   16125
      _ExtentX        =   28443
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lvListOK1 
      Height          =   1815
      Left            =   180
      TabIndex        =   12
      Top             =   4725
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   3201
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
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":"
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
      Index           =   4
      Left            =   1575
      TabIndex        =   43
      Top             =   900
      Width           =   150
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   ":"
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
      Index           =   3
      Left            =   1575
      TabIndex        =   42
      Top             =   450
      Width           =   150
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Product"
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
      Index           =   1
      Left            =   135
      TabIndex        =   40
      Top             =   900
      Width           =   1455
   End
   Begin PRD.ucChartBar ucChartBar1 
      Height          =   3030
      Left            =   180
      TabIndex        =   39
      Top             =   6975
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   5345
      Title           =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LegendAlign     =   0
      LegendVisible   =   0   'False
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleForeColor  =   0
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808080&
      Caption         =   "INFORMATION PRODUCT OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   180
      TabIndex        =   38
      Top             =   4410
      Width           =   16080
   End
   Begin VB.Label LblProd 
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   8
      Left            =   14985
      TabIndex        =   37
      Top             =   1305
      Width           =   1275
   End
   Begin VB.Label LblProd 
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   7
      Left            =   11970
      TabIndex        =   36
      Top             =   1710
      Width           =   1275
   End
   Begin VB.Label LblProd 
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   6
      Left            =   11970
      TabIndex        =   35
      Top             =   1305
      Width           =   1275
   End
   Begin VB.Label LblProd 
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   14985
      TabIndex        =   34
      Top             =   1710
      Width           =   1275
   End
   Begin VB.Label LblProd 
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   5940
      TabIndex        =   33
      Top             =   1755
      Width           =   3930
   End
   Begin VB.Label LblProd 
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   5940
      TabIndex        =   32
      Top             =   1350
      Width           =   3930
   End
   Begin VB.Label LblProd 
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   1755
      TabIndex        =   31
      Top             =   1755
      Width           =   2985
   End
   Begin VB.Label LblProd 
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   1755
      TabIndex        =   30
      Top             =   1350
      Width           =   2985
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "TARGET YIELD [%]"
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
      Height          =   240
      Left            =   10350
      TabIndex        =   29
      Top             =   5760
      Width           =   1950
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   " [%] TARGET"
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
      Height          =   240
      Left            =   14445
      TabIndex        =   28
      Top             =   5715
      Width           =   1770
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      Caption         =   "PROD YIELD [%]"
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
      Height          =   240
      Left            =   12510
      TabIndex        =   27
      Top             =   5715
      Width           =   1770
   End
   Begin VB.Label Label32 
      BackColor       =   &H00808080&
      Caption         =   "TOP 10 DATA IDLE TIME"
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
      Left            =   8550
      TabIndex        =   26
      Top             =   6615
      Width           =   7710
   End
   Begin VB.Label Label31 
      BackColor       =   &H00808080&
      Caption         =   "TOP 10 DATA NG"
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
      Left            =   180
      TabIndex        =   25
      Top             =   6615
      Width           =   8160
   End
   Begin VB.Label Label27 
      BackColor       =   &H00808080&
      Caption         =   "INFORMATION COUNTER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   180
      TabIndex        =   24
      Top             =   2295
      Width           =   16080
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Berat Runner"
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
      Left            =   13635
      TabIndex        =   23
      Top             =   1395
      Width           =   1545
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Part No"
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
      Left            =   180
      TabIndex        =   22
      Top             =   1800
      Width           =   1545
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Internal Part ID"
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
      Left            =   180
      TabIndex        =   21
      Top             =   1395
      Width           =   1500
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "GROSS PRODUCT"
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
      Height          =   240
      Left            =   10350
      TabIndex        =   20
      Top             =   4725
      Width           =   1950
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "OK PRODUCT"
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
      Height          =   240
      Left            =   14445
      TabIndex        =   19
      Top             =   4725
      Width           =   1770
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "NG PRODUCT"
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
      Height          =   240
      Left            =   12510
      TabIndex        =   18
      Top             =   4725
      Width           =   1725
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
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
      Height          =   330
      Index           =   5
      Left            =   5085
      TabIndex        =   17
      Top             =   1800
      Width           =   1005
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Warna"
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
      Left            =   5085
      TabIndex        =   16
      Top             =   1395
      Width           =   1230
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Cavity"
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
      Left            =   10350
      TabIndex        =   15
      Top             =   1395
      Width           =   1545
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Production Yield"
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
      Left            =   10350
      TabIndex        =   14
      Top             =   1800
      Width           =   1545
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Berat Produk"
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
      Left            =   13635
      TabIndex        =   13
      Top             =   1845
      Width           =   1545
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Customer"
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
      Index           =   2
      Left            =   135
      TabIndex        =   2
      Top             =   450
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00FFFFFF&
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
      Height          =   330
      Index           =   0
      Left            =   1710
      TabIndex        =   1
      Top             =   450
      Width           =   8115
   End
End
Attribute VB_Name = "FrmDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim srcItem                         As ListItem
Dim srcRecord                       As String
Dim srcProduct                      As Variant
Dim sSQL                            As String
Dim sHour                           As String

Dim miliseconds As Integer, seconds As Integer, minutes As Integer, hours As Integer

Dim xTime As Integer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cboProduct_Click()

    If CboProduct.ListIndex = 0 Then
        Call Get_Data(p_eng_product_1)
    ElseIf CboProduct.ListIndex = 1 Then
        Call Get_Data(p_eng_product_2)
    ElseIf CboProduct.ListIndex = 2 Then
        Call Get_Data(p_eng_product_3)
    ElseIf CboProduct.ListIndex = 3 Then
        Call Get_Data(p_eng_product_4)
    End If
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler

    ShowList1 lvList

    ShowList2 LvListIdle1

    ShowListOK lvListOK1

    If p_status_prod_1 = True Then
        CboProduct.AddItem p_prod_name_1
    End If

    If p_status_prod_2 = True Then
        CboProduct.AddItem p_prod_name_2
    End If

    If p_status_prod_3 = True Then
        CboProduct.AddItem p_prod_name_3
    End If

    If p_status_prod_4 = True Then
        CboProduct.AddItem p_prod_name_4
    End If
    
    CboProduct.ListIndex = 0
    
    Call Get_Data(p_eng_product_1)
    


Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub GetProd(iID As String, tbox0 As Label, tbox1 As Label, tbox2 As Label, tbox3 As Label, _
                     tbox4 As Label, tbox5 As Label, tbox6 As Label, tbox7 As Label)
On Error GoTo ErrHandler

    Dim Rs As New Recordset
    Dim recSQL As String
    
    Rs.CursorLocation = adUseClient
    
    recSQL = "SELECT a.id,a.cycle_time_ia,a.internal_part_id,a.customer_part_number,a.customer_part_name,"
    recSQL = recSQL & " a.eng_material_id,a.eng_color_id,a.cavity,a.prod_yield,"
    recSQL = recSQL & " a.weight_gr,a.weight_runner_gr, b.name as color_name, c.name as material_name"
    recSQL = recSQL & " from sip_production.eng_products a"
    recSQL = recSQL & " left join sip_production.eng_colors b on a.eng_color_id = b.id"
    recSQL = recSQL & " left join sip_production.eng_materials c on a.eng_material_id = c.id "
    recSQL = recSQL & " where a.status_plant_3 = 'active' and a.id = " & iID & ""
    
    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open recSQL, CN, adOpenStatic, adLockOptimistic
    
    If Rs.RecordCount > 0 Then
        tbox0.Caption = Rs.Fields("internal_part_id")
        tbox1.Caption = Rs.Fields("customer_part_number")
        tbox2.Caption = IIf(IsNull(Rs.Fields("color_name")), "", Rs.Fields("color_name"))
        tbox3.Caption = IIf(IsNull(Rs.Fields("material_name")), "", Rs.Fields("material_name"))
        tbox4.Caption = IIf(IsNull(Rs.Fields("weight_gr")), "", Rs.Fields("weight_gr"))
        tbox5.Caption = IIf(IsNull(Rs.Fields("cavity")), "", Rs.Fields("cavity")) 'RS.Fields("cavity")
        tbox6.Caption = IIf(IsNull(Rs.Fields("prod_yield")), "", Rs.Fields("prod_yield")) 'RS.Fields("prod_yield")
        tbox7.Caption = IIf(IsNull(Rs.Fields("weight_runner_gr")), "", Rs.Fields("weight_runner_gr")) 'RS.Fields("weight_runner_gr")
    End If

    Set Rs = Nothing

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
    
End Sub



Private Sub Form_Resize()
On Error Resume Next

    If WindowState <> vbMinimized Then
        b8TitleBar1.Width = Me.ScaleWidth
        
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500

    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MAIN.RemoveChild Me.Name
    Set FrmDashboard = Nothing
End Sub

Private Sub Form_Activate()
On Error Resume Next

With MAIN
    Me.BackColor = .ACPMenu.BackColor
    Me.Picture = .ACPMenu.LoadBackground
End With
MAIN.ActivateChild Me


End Sub

Private Sub FillListview(sProd As String, List As listview)
On Error Resume Next

    Dim Rs As New ADODB.Recordset
    Dim sSQL As Variant

    Rs.CursorLocation = adUseClient

 sSQL = "select plant_mark,prod_machine_id,mkt_customer_id,eng_product_id,period_shift, STATUS,"
 sSQL = sSQL & " C08,C09,C10,C11,C12,C13,C14,C15,(C08+C09+C10+C11+C12+C13+C14+C15) as Shift_1,"
 sSQL = sSQL & " C16,C17,C18,C19,C20,C21,C22,C23,(C16+C17+C18+C19+C20+C21+C22+C23) as Shift_2,"
 sSQL = sSQL & " C00,C01,C02,C03,C04,C05,C06,C07,(C00+C01+C02+C03+C04+C05+C06+C07) as Shift_3"
 sSQL = sSQL & " From (select plant_mark,prod_machine_id,mkt_customer_id,eng_product_id,period_shift, STATUS,"
 sSQL = sSQL & " SUM(IF(period_hour = '08', jumlah, 0)) AS 'C08',"
 sSQL = sSQL & " SUM(IF(period_hour = '09', jumlah, 0)) AS 'C09',"
 sSQL = sSQL & " SUM(IF(period_hour = '10', jumlah, 0)) AS 'C10',"
 sSQL = sSQL & " SUM(IF(period_hour = '11', jumlah, 0)) AS 'C11',"
 sSQL = sSQL & " SUM(IF(period_hour = '12', jumlah, 0)) AS 'C12',"
 sSQL = sSQL & " SUM(IF(period_hour = '13', jumlah, 0)) AS 'C13',"
 sSQL = sSQL & " SUM(IF(period_hour = '14', jumlah, 0)) AS 'C14',"
 sSQL = sSQL & " SUM(IF(period_hour = '15', jumlah, 0)) AS 'C15',"
 sSQL = sSQL & " SUM(IF(period_hour = '16', jumlah, 0)) AS 'C16',"
 sSQL = sSQL & " SUM(IF(period_hour = '17', jumlah, 0)) AS 'C17',"
 sSQL = sSQL & " SUM(IF(period_hour = '18', jumlah, 0)) AS 'C18',"
 sSQL = sSQL & " SUM(IF(period_hour = '19', jumlah, 0)) AS 'C19',"
 sSQL = sSQL & " SUM(IF(period_hour = '20', jumlah, 0)) AS 'C20',"
 sSQL = sSQL & " SUM(IF(period_hour = '21', jumlah, 0)) AS 'C21',"
 sSQL = sSQL & " SUM(IF(period_hour = '22', jumlah, 0)) AS 'C22',"
 sSQL = sSQL & " SUM(IF(period_hour = '23', jumlah, 0)) AS 'C23',"
 sSQL = sSQL & " SUM(IF(period_hour = '00', jumlah, 0)) AS 'C00',"
 sSQL = sSQL & " SUM(IF(period_hour = '01', jumlah, 0)) AS 'C01',"
 sSQL = sSQL & " SUM(IF(period_hour = '02', jumlah, 0)) AS 'C02',"
 sSQL = sSQL & " SUM(IF(period_hour = '03', jumlah, 0)) AS 'C03',"
 sSQL = sSQL & " SUM(IF(period_hour = '04', jumlah, 0)) AS 'C04',"
 sSQL = sSQL & " SUM(IF(period_hour = '05', jumlah, 0)) AS 'C05',"
 sSQL = sSQL & " SUM(IF(period_hour = '06', jumlah, 0)) AS 'C06',"
 sSQL = sSQL & " SUM(IF(period_hour = '07', jumlah, 0)) AS 'C07'"
 sSQL = sSQL & " From (select a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift,a.period_hour,"
            sSQL = sSQL & " sum(a.counter_ok) as jumlah,'1. SHOT' as status from sip_production.prod_runnings a group by a.plant_mark,"
            sSQL = sSQL & " a.prod_machine_id , a.mkt_customer_id, a.eng_product_id, a.period_shift, a.period_hour"
sSQL = sSQL & " Union select a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift,period_hour,"
             sSQL = sSQL & " sum(counter_ng) AS jumlah,'3. NG PROD' as status from sip_production.prod_data_ngs a"
             sSQL = sSQL & " where a.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "' and status = 'active' group by a.plant_mark,a.prod_machine_id, a.mkt_customer_id,a.eng_product_id,a.period_shift,period_hour"
sSQL = sSQL & " Union select a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift,"
             sSQL = sSQL & " a.period_hour, sum(a.counter_ok * b.cavity) as jumlah,'2. GROSS' as status from sip_production.prod_runnings a"
             sSQL = sSQL & " inner join sip_production.eng_products b on a.eng_product_id = b.id group by a.plant_mark,a.prod_machine_id,"
             sSQL = sSQL & " a.mkt_customer_id , a.eng_product_id, a.period_shift, a.period_hour"
sSQL = sSQL & " Union select xx.plant_mark,xx.prod_machine_id,xx.mkt_customer_id,xx.eng_product_id,xx.period_shift,xx.period_hour,"
            sSQL = sSQL & " (xx.gross_produksi-ifnull(yy.ng,0)) as net_produksi,'4. NET_PROD' as status"
            sSQL = sSQL & " from (select a.plant_mark,a.prod_machine_id,a.mkt_customer_id, a.eng_product_id,a.period_shift,a.period_hour, a.counter_ok as shot,"
                            sSQL = sSQL & " sum(a.counter_ok * b.cavity) as gross_produksi from sip_production.prod_runnings a"
                            sSQL = sSQL & " inner join sip_production.eng_products b on a.eng_product_id = b.id"
                            sSQL = sSQL & " GROUP by a.plant_mark,a.prod_machine_id,a.mkt_customer_id, a.eng_product_id,a.period_shift,a.period_hour) xx"
            sSQL = sSQL & " left join (select a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift,"
                            sSQL = sSQL & " period_hour, sum(counter_ng) as ng from sip_production.prod_data_ngs a"
                            sSQL = sSQL & " where a.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "' and status = 'active' group by a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift,period_hour) yy"
            sSQL = sSQL & " on xx.plant_mark = yy.plant_mark and xx.prod_machine_id = yy.prod_machine_id and xx.mkt_customer_id = yy.mkt_customer_id"
            sSQL = sSQL & " and xx.eng_product_id = yy.eng_product_id and xx.period_shift = yy.period_shift and xx.period_hour = yy.period_hour ) as x"
    sSQL = sSQL & " where x.plant_mark = '" & p_plant_mark & "'"
    sSQL = sSQL & " and x.prod_machine_id = '" & p_prod_machine_id & "'"
    sSQL = sSQL & " and x.mkt_customer_id = '" & p_mkt_customer_id & "'"
    sSQL = sSQL & " and x.eng_product_id = '" & sProd & "'"
    sSQL = sSQL & " and x.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "'"
sSQL = sSQL & " group by plant_mark,prod_machine_id,mkt_customer_id,eng_product_id,period_shift, STATUS) data_produksi"


    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
    With List
    
        .ListItems.Clear
        Do While Not Rs.EOF
        Set srcItem = .ListItems.Add(, , Rs.Fields("Status"), 1, 1)
            srcItem.SubItems(1) = Rs.Fields("C08")
            srcItem.SubItems(2) = Rs.Fields("C09")
            srcItem.SubItems(3) = Rs.Fields("C10")
            srcItem.SubItems(4) = Rs.Fields("C11")
            srcItem.SubItems(5) = Rs.Fields("C12")
            srcItem.SubItems(6) = Rs.Fields("C13")
            srcItem.SubItems(7) = Rs.Fields("C14")
            srcItem.SubItems(8) = Rs.Fields("C15")
            srcItem.SubItems(9) = Rs.Fields("Shift_1")
            srcItem.SubItems(10) = Rs.Fields("C16")
            srcItem.SubItems(11) = Rs.Fields("C17")
            srcItem.SubItems(12) = Rs.Fields("C18")
            srcItem.SubItems(13) = Rs.Fields("C19")
            srcItem.SubItems(14) = Rs.Fields("C20")
            srcItem.SubItems(15) = Rs.Fields("C21")
            srcItem.SubItems(16) = Rs.Fields("C22")
            srcItem.SubItems(17) = Rs.Fields("C23")
            srcItem.SubItems(18) = Rs.Fields("Shift_2")
            srcItem.SubItems(19) = Rs.Fields("C00")
            srcItem.SubItems(20) = Rs.Fields("C01")
            srcItem.SubItems(21) = Rs.Fields("C02")
            srcItem.SubItems(22) = Rs.Fields("C03")
            srcItem.SubItems(23) = Rs.Fields("C04")
            srcItem.SubItems(24) = Rs.Fields("C05")
            srcItem.SubItems(25) = Rs.Fields("C06")
            srcItem.SubItems(26) = Rs.Fields("C07")
            srcItem.SubItems(27) = Rs.Fields("Shift_3")
            Rs.MoveNext
        Loop
    End With
    
    Rs.Close
    Set Rs = Nothing

'Exit Sub
'ErrHandler:
'MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
'

End Sub

Private Sub ShowList1(List As listview)

    With List
        .GridLines = True
        .View = lvwReport
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "INFO", 1500
        .ColumnHeaders.Add , , "08", 500
        .ColumnHeaders.Add , , "09", 500
        .ColumnHeaders.Add , , "10", 500
        .ColumnHeaders.Add , , "11", 500
        .ColumnHeaders.Add , , "12", 500
        .ColumnHeaders.Add , , "13", 500
        .ColumnHeaders.Add , , "14", 500
        .ColumnHeaders.Add , , "15", 500
        .ColumnHeaders.Add , , "SHIFT-1", 900
        .ColumnHeaders.Add , , "16", 500
        .ColumnHeaders.Add , , "17", 500
        .ColumnHeaders.Add , , "18", 500
        .ColumnHeaders.Add , , "19", 500
        .ColumnHeaders.Add , , "20", 500
        .ColumnHeaders.Add , , "21", 500
        .ColumnHeaders.Add , , "22", 500
        .ColumnHeaders.Add , , "23", 500
        .ColumnHeaders.Add , , "SHIFT-2", 900
        .ColumnHeaders.Add , , "00", 500
        .ColumnHeaders.Add , , "01", 500
        .ColumnHeaders.Add , , "02", 500
        .ColumnHeaders.Add , , "03", 500
        .ColumnHeaders.Add , , "04", 500
        .ColumnHeaders.Add , , "05", 500
        .ColumnHeaders.Add , , "06", 500
        .ColumnHeaders.Add , , "07", 500
        .ColumnHeaders.Add , , "SHIFT-3", 900
        .ListItems.Clear
                
    End With
End Sub

Private Sub ShowList2(List As listview)

    With List
        .GridLines = True
        .View = lvwReport
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "START TIME", 1500
        .ColumnHeaders.Add , , "IDLE TIME", 1500
        .ColumnHeaders.Add , , "IDLE NAME", 2500
        .ColumnHeaders.Add , , "USER", 2200
        .ListItems.Clear
    End With
    
End Sub


Private Sub ShowListOK(List As listview)

    With List
        .GridLines = True
        .View = lvwReport
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "NO", 700
        .ColumnHeaders.Add , , "PRODUCT", 3500
        .ColumnHeaders.Add , , "PERIODE", 1300
        .ColumnHeaders.Add , , "SHIFT", 1000
        .ColumnHeaders.Add , , "OK [FULL BOX]", 1300
        .ColumnHeaders.Add , , "SISA", 1000
        .ColumnHeaders.Add , , "HOLD", 1000
        .ListItems.Clear
    End With
    
End Sub


Private Sub ChartData(sProd As String, ChatBar As ucChartBar)
On Error Resume Next
    Dim Value As Collection
    Dim i As Long, j As Long
    Dim Palette() As String
    Dim Users() As String
    Dim MyArray() As String
    Dim Rs As New ADODB.Recordset
    Dim sSQL As Variant

    sSQL = "SELECT b.name as product_name,c.name as ng_name,SUM(a.counter_ng) AS total"
    sSQL = sSQL & " FROM sip_production.prod_data_ngs a"
    sSQL = sSQL & " LEFT JOIN sip_production.eng_products b ON a.eng_product_id = b.id"
    sSQL = sSQL & " LEFT JOIN sip_production.prod_ngs c ON a.prod_ng_id = c.id"
    sSQL = sSQL & " where a.plant_mark = '" & p_plant_mark & "'"
    sSQL = sSQL & " and a.prod_machine_id = '" & p_prod_machine_id & "'"
    sSQL = sSQL & " and a.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "'"
    sSQL = sSQL & " and a.eng_product_id = '" & sProd & "'"
    sSQL = sSQL & " GROUP BY a.prod_machine_id,a.prod_ng_id order by total desc limit 10"



    Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
    
    
    ChatBar.Clear

    i = 0
    If Rs.RecordCount > 0 Then
        Do While Not Rs.EOF
            ReDim Preserve Users(i)
            ReDim Preserve MyArray(i)
            
            Users(i) = Rs.Fields("ng_name")
            MyArray(i) = Rs.Fields("total")
            
            Rs.MoveNext
            i = i + 1
        Loop
    
        Palette = Split("&HFF8D11,&HA744E0,&H376CE6,&H40AB1A,&H7B006B,&H7B006B,&H7B006B", ",")
    
        Set Value = New Collection
        For i = 0 To UBound(Users)
            Value.Add Users(i)
        Next
        
        ChatBar.AddAxisItems Value, False, 0, ccEnter
    
    
        Set Value = New Collection
        For j = 0 To UBound(MyArray)
            Value.Add MyArray(j)
        Next
    
        ChatBar.AddSerie "Leader", CLng(Palette(1)), Value
        ChatBar.LabelsVisible = True
        ChatBar.Refresh
    End If
    
End Sub

Private Sub ListIdletime(List As listview)
On Error GoTo ErrHandler
    Dim Rs As New Recordset
    Dim sSQL As String

    Rs.CursorLocation = adUseClient
    sSQL = "SELECT a.hrd_employee_id,a.plant_mark,a.prod_machine_id,a.prod_idletime_id,"
    sSQL = sSQL & " a.period_shift,a.start_idle,a.end_idle,a.idle_time ,c.name as idle_name,c.description,d.name as karyawan"
    sSQL = sSQL & " from sip_production.prod_machine_idles a"
    sSQL = sSQL & " INNER JOIN sip_production.prod_idletimes c ON a.prod_idletime_id = c.id"
    sSQL = sSQL & " LEFT JOIN sip_production.hrd_employees d ON a.hrd_employee_id = d.id "
    sSQL = sSQL & " where a.plant_mark = '" & p_plant_mark & "'"
    sSQL = sSQL & " and a.prod_machine_id = '" & p_prod_machine_id & "'"
    sSQL = sSQL & " AND date(a.period_shift) = '" & Format(Date, "yyyy-mm-dd") & "' ORDER BY a.idle_time DESC limit 10"

    Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic

    With List

        .ListItems.Clear
        
        Do While Not Rs.EOF
        Set srcItem = .ListItems.Add(, , Format(Rs.Fields("start_idle"), "hh:mm:ss"), 1, 1)
            srcItem.SubItems(1) = IIf(IsNull(Format(Rs.Fields("idle_time"), "hh:mm:ss")), "", Format(Rs.Fields("idle_time"), "hh:mm:ss"))
            srcItem.SubItems(2) = Rs.Fields("idle_name")
            srcItem.SubItems(3) = IIf(IsNull(Rs.Fields("karyawan")), "", Rs.Fields("karyawan"))
    
            Rs.MoveNext
        Loop
    End With
    
    Set Rs = Nothing

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
    
End Sub


Private Sub Summary(sProd As String, tGross As TextBox, tNG As TextBox, tOK As TextBox, tTargetYield As TextBox, tTotalYield As TextBox)
On Error GoTo ErrHandler

    Dim Rs As New Recordset
    Dim sSQL As String

    Rs.CursorLocation = adUseClient
    
    sSQL = "SELECT a.plant_mark,a.prod_machine_id, c.number as machine_no,"
    sSQL = sSQL & " a.eng_product_id,"
    sSQL = sSQL & " b.cavity,b.prod_yield AS target_yield,b.weight_gr,b.cycle_time_ia,"
    sSQL = sSQL & " (floor(3600/b.cycle_time_ia) * count(a.period_hour))  AS target_shot, sum(a.counter_ok) jumlah_shot,"
    sSQL = sSQL & " sum(a.counter_ok) * b.cavity AS gross_produksi, ifnull(data_ngs.jumlah_ng,0) AS total_ng ,"
    sSQL = sSQL & " ifnull(((sum(a.counter_ok) * b.cavity) - ifnull(data_ngs.jumlah_ng,0)),0) AS net_produksi,"
    sSQL = sSQL & " ifnull(ROUND((((sum(a.counter_ok) * b.cavity) - ifnull(data_ngs.jumlah_ng,0)) / (sum(a.counter_ok) * b.cavity)) * 100,2),0) AS prod_yield"
    sSQL = sSQL & " FROM prod_runnings a"
    sSQL = sSQL & " LEFT JOIN eng_products b ON a.eng_product_id = b.id"
    sSQL = sSQL & " LEFT JOIN (SELECT d.plant_mark,d.prod_machine_id,d.eng_product_id,d.period_shift, sum(d.counter_ng) jumlah_ng FROM prod_data_ngs d"
                sSQL = sSQL & " where d.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "'"
                sSQL = sSQL & " GROUP BY d.plant_mark,d.prod_machine_id,d.eng_product_id,d.period_shift) AS data_ngs ON a.plant_mark = data_ngs.plant_mark"
                sSQL = sSQL & " AND a.prod_machine_id = data_ngs.prod_machine_id AND a.eng_product_id = data_ngs.eng_product_id"
                sSQL = sSQL & " AND a.period_shift = data_ngs.period_shift"
    sSQL = sSQL & " LEFT JOIN prod_machines c ON a.prod_machine_id = c.id"
    sSQL = sSQL & " where a.plant_mark = '" & p_plant_mark & "' and"
    sSQL = sSQL & " a.prod_machine_id = '" & p_prod_machine_id & "' and"
    sSQL = sSQL & " a.eng_product_id = '" & sProd & "' and"
    sSQL = sSQL & " a.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "'"
    sSQL = sSQL & " GROUP BY a.plant_mark,a.eng_product_id,a.prod_machine_id"

    Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
    
    If Rs.RecordCount > 0 Then
        tGross.text = IIf(IsNull(Rs.Fields("gross_produksi")), "0", Rs.Fields("gross_produksi"))
        tNG.text = IIf(IsNull(Rs.Fields("total_ng")), "0", Rs.Fields("total_ng"))
        tOK.text = IIf(IsNull(Rs.Fields("net_produksi")), "0", Rs.Fields("net_produksi"))
        
        tTargetYield.text = IIf(IsNull(Rs.Fields("target_yield")), "0", Rs.Fields("target_yield"))
        tTotalYield.text = IIf(IsNull(Rs.Fields("prod_yield")), "0", Rs.Fields("prod_yield"))

    End If
    Set Rs = Nothing

    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
    
    
End Sub



Private Sub Timer1_Timer()
    ProgressBar1.Max = 120
    xTime = xTime + 1
    If xTime >= 120 Then
        If CboProduct.ListIndex = 0 Then
            Call Get_Data(p_eng_product_1)
        ElseIf CboProduct.ListIndex = 1 Then
            Call Get_Data(p_eng_product_2)
        ElseIf CboProduct.ListIndex = 2 Then
            Call Get_Data(p_eng_product_3)
        ElseIf CboProduct.ListIndex = 3 Then
            Call Get_Data(p_eng_product_4)
        End If
        xTime = 0
        ProgressBar1.Value = 0
    Else
        ProgressBar1.Value = ProgressBar1.Value + 1
    End If
End Sub


Private Sub Get_Data(ByVal sProduct As String)

On Error GoTo ErrHandler

    GetProd sProduct, LblProd(1), LblProd(2), LblProd(3), LblProd(4), LblProd(5), LblProd(6), LblProd(7), LblProd(8)
    lblInfo(0).Caption = p_customer_name
    'LblProd(1).Caption = p_int_part_3
    
    FillListview sProduct, lvList

    ChartData sProduct, ucChartBar1

    ListOK sProduct, lvListOK1

    ListIdletime LvListIdle1

    Summary sProduct, txtgross_1, txtNg_1, txtOk_1, txtTarget_Yield_1, txtTotalYield_1

    Persen_Target sProduct, txtPercTarget_1

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
  
End Sub

Private Sub ListOK(sProd As String, List As listview)
On Error GoTo ErrHandler
    Dim i As Integer
    Dim Rs As New ADODB.Recordset
    Dim sSQL As Variant

    i = 1

sSQL = "select a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,b.name as product_name, a.period_shift,a.shift,"
sSQL = sSQL & " SUM(IF(product_status = 'ok', qty, 0) ) AS  'ok',"
sSQL = sSQL & " SUM(IF(product_status = 'sisa', qty, 0) ) AS  'sisa',"
sSQL = sSQL & " SUM(IF(product_status = 'hold', qty, 0) ) AS  'hold'"
sSQL = sSQL & " from prod_result_logs a"
sSQL = sSQL & " left join eng_products b on a.eng_product_id = b.id"
    sSQL = sSQL & " where a.`status` = 'active'"
    sSQL = sSQL & " and a.plant_mark = '" & p_plant_mark & "'"
    sSQL = sSQL & " and a.prod_machine_id = '" & p_prod_machine_id & "'"
    sSQL = sSQL & " and a.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "'"
    sSQL = sSQL & " and a.eng_product_id = '" & sProd & "'"
sSQL = sSQL & " group by a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift,a.shift"

    Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
    
    With List

        .ListItems.Clear
        Do While Not Rs.EOF
        Set srcItem = .ListItems.Add(, , i, 1, 1)
            srcItem.SubItems(1) = Rs.Fields("product_name")
            srcItem.SubItems(2) = Format(Rs.Fields("period_shift"), "yyyy-mm-dd")
            srcItem.SubItems(3) = Rs.Fields("shift")
            srcItem.SubItems(4) = Rs.Fields("ok")
            srcItem.SubItems(5) = Rs.Fields("sisa")
            srcItem.SubItems(6) = Rs.Fields("hold")
            
            i = i + 1
            Rs.MoveNext
        Loop
    End With
    
    Set Rs = Nothing


Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub


Private Sub Persen_Target(sProd As String, tPercent As TextBox)
On Error GoTo ErrHandler

    Dim Rs As New Recordset
    Dim sSQL As String
    Dim sTime As String
    Rs.CursorLocation = adUseClient
    
    Dim sPeriod_hour As Variant
    Dim sIdle_time As Variant
    Dim sCounter As Variant
    Dim sRunning_hour As Variant
    Dim sTarget_CT As Variant

    'lblRunning_hours.Caption = Val(DateDiff("s", "08:00:00", Format(Now, "hh:mm:ss"))) / p_cycle_time_1

    sSQL = "SELECT a.plant_mark,a.prod_machine_id,a.eng_product_id,"
    sSQL = sSQL & " a.period_shift, sum(a.counter_ok) as total_shot,ifnull(c.t_sec_idle,0) as  t_sec_idle,a.period_hour"
    sSQL = sSQL & " FROM sip_production.prod_runnings a"
    sSQL = sSQL & " LEFT JOIN  (select plant_mark,prod_machine_id, period_shift,sum(time_to_sec(idle_time)) AS t_sec_idle"
                    sSQL = sSQL & " from prod_machine_idles where eng_product_1 = '" & p_eng_product_1 & "' "
                    sSQL = sSQL & " group by plant_mark,prod_machine_id,mkt_customer_id,period_shift) as c"
                    sSQL = sSQL & " ON a.plant_mark = c.plant_mark and a.prod_machine_id = c.prod_machine_id"
                    sSQL = sSQL & " and a.period_shift = c.period_shift"
    sSQL = sSQL & " where a.plant_mark = '" & p_plant_mark & "' and"
    sSQL = sSQL & " a.prod_machine_id = '" & p_prod_machine_id & "' and"
    sSQL = sSQL & " a.eng_product_id = '" & sProd & "' and"
    sSQL = sSQL & " a.period_shift = '" & Format(p_shift, "yyyy-mm-dd") & "'"
    sSQL = sSQL & " Group by a.plant_mark,a.prod_machine_id,a.eng_product_id,a.period_shift order by a.period_hour asc"

    Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic

    If Rs.RecordCount > 0 Then
        'tPercent.Text = Round(IIf(IsNull(Rs.Fields("persen_target")), "0", Rs.Fields("persen_target")), 0)
        sPeriod_hour = Format(Rs.Fields("period_shift"), "yyyy-mm-dd") & " " & Rs.Fields("period_hour") & ":00:00"
        sIdle_time = Rs.Fields("t_sec_idle")
        sCounter = Rs.Fields("total_shot")
        sTarget_CT = Round(3600 / p_cycle_time_1, 2)
        
        sRunning_hour = Round(Val(DateDiff("s", sPeriod_hour, Format(Now, "yyyy-mm-dd hh:mm:ss")) - sIdle_time) / 3600, 2)
        'Rumus Persen target
        'Jumlah Shot / (running_hour - idle_time) / CT
        
        tPercent.text = Round((sCounter / sTarget_CT) / sRunning_hour * 100, 2)
    End If
    Set Rs = Nothing


Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation


End Sub


