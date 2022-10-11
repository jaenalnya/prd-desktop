VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LvButton.ocx"
Object = "{9EDDC69F-10E8-4DE7-9648-EC8A45F005C0}#1.0#0"; "B8CONT~2.OCX"
Begin VB.Form frmProduction 
   Caption         =   "Production"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18525
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmProduction.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   18525
   Begin b8Controls4.b8TitleBar b8TitleBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   13515
      _ExtentX        =   23839
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
   Begin VB.Frame Frame1 
      Height          =   5025
      Left            =   90
      TabIndex        =   3
      Top             =   945
      Width           =   8895
      Begin VB.PictureBox PicWarning 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1770
         Left            =   135
         ScaleHeight     =   1740
         ScaleWidth      =   5475
         TabIndex        =   7
         Top             =   225
         Visible         =   0   'False
         Width           =   5505
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "FILE WI / CCP / PS TIDAK TERSEDIA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1410
            Left            =   315
            TabIndex        =   8
            Top             =   315
            Width           =   4740
         End
      End
   End
   Begin lvButton.lvButtons_H cmdPS 
      Height          =   435
      Left            =   4725
      TabIndex        =   0
      ToolTipText     =   "Packing Standard"
      Top             =   450
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   767
      Caption         =   "Packing Standard [1]"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      cBhover         =   33023
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmProduction.frx":617A
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdWI 
      Height          =   435
      Left            =   90
      TabIndex        =   1
      ToolTipText     =   "Work Instruction"
      Top             =   450
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   767
      Caption         =   "Work Instruction [1]"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      cBhover         =   33023
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmProduction.frx":C304
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdCP 
      Height          =   435
      Left            =   9315
      TabIndex        =   2
      ToolTipText     =   "Critical Point Check"
      Top             =   450
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   767
      Caption         =   "Critical Point [1]"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      cBhover         =   33023
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmProduction.frx":1248E
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdPS2 
      Height          =   435
      Left            =   6885
      TabIndex        =   4
      ToolTipText     =   "Packing Standard"
      Top             =   450
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   767
      Caption         =   "Packing Standard [2]"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      cBhover         =   33023
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmProduction.frx":18618
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdCP2 
      Height          =   435
      Left            =   11475
      TabIndex        =   5
      ToolTipText     =   "Critical Point Check"
      Top             =   450
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   767
      Caption         =   "Critical Point [2]"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      cBhover         =   33023
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmProduction.frx":1E7A2
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdWI2 
      Height          =   435
      Left            =   2250
      TabIndex        =   6
      ToolTipText     =   "Work Instruction"
      Top             =   450
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   767
      Caption         =   "Work Instruction [2]"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      cBhover         =   33023
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmProduction.frx":2492C
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H cmdUM1 
      Height          =   435
      Left            =   13905
      TabIndex        =   10
      ToolTipText     =   "Critical Point Check"
      Top             =   450
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   767
      Caption         =   "User Manual"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   0
      cBhover         =   33023
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "frmProduction.frx":2AAB6
      cBack           =   16777215
   End
End
Attribute VB_Name = "frmProduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_objPDF As AcroPDFLibCtl.AcroPDF 'Declare an object of type AcroPDF
Private m_strFilePath As String 'Declare a string for the PDF Filename and Path
Dim izoomWI As Integer
Dim izoomPS As Integer
Dim izoomCCP As Integer

Private Sub cmdCP_Click()
On Error GoTo ErrHandler

    LoadPDF App.Path & "\Picture\CCP\" & p_int_part_1 & ".pdf", izoomCCP
    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub cmdCP2_Click()
On Error GoTo ErrHandler

    LoadPDF App.Path & "\Picture\CCP\" & p_int_part_2 & ".pdf", izoomCCP

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub



Private Sub cmdPS_Click()
On Error GoTo ErrHandler

    LoadPDF App.Path & "\Picture\PS\" & p_int_part_1 & ".pdf", izoomPS


Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub LoadPDF(sPath As String, iZoom As Integer)
On Error GoTo ErrHandler

    If Dir(sPath) <> "" Then
        PicWarning.Visible = False
        
        m_objPDF.LoadFile sPath
        m_objPDF.setShowToolbar False
        m_objPDF.setLayoutMode "SinglePage"
        m_objPDF.setPageMode "none"
        
        'Set the Zoom view according to the value specified. ranges from 0 and onwards
        m_objPDF.setZoom iZoom
        
        'Move and Resize the object in relation to its container/form
        With m_objPDF
           '.Move 125, 175, Me.Width, Me.Height 'x-position, y-position, width, height
           .Move 125, 175, Frame1.Width - 300, Frame1.Height - 300 'x-position, y-position, width, height
        End With
        
        m_objPDF.Visible = True
    Else
        PicWarning.Visible = True
    End If
    


    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub cmdPS2_Click()
On Error GoTo ErrHandler

    LoadPDF App.Path & "\Picture\PS\" & p_int_part_2 & ".pdf", izoomPS
    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub cmdUM1_Click()
On Error GoTo ErrHandler

    LoadPDF App.Path & "\User Manual.pdf", izoomWI

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub cmdWI_Click()
On Error GoTo ErrHandler

    LoadPDF App.Path & "\Picture\WI\" & p_int_part_1 & ".pdf", izoomWI

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub cmdWI2_Click()
On Error GoTo ErrHandler
    LoadPDF App.Path & "\Picture\WI\" & p_int_part_2 & ".pdf", izoomWI
    
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler

izoomPS = ReadINI("SETTING", "PS", App.Path & "\Settings.ini")
izoomWI = ReadINI("SETTING", "WI", App.Path & "\Settings.ini")
izoomCCP = ReadINI("SETTING", "CCP", App.Path & "\Settings.ini")

   Set m_objPDF = Controls.Add("AcroPDF.PDF.1", "AcroPDF1") 'This will add the PDF Browser control to the form on runtime. The "Test" is the control's name
   Set m_objPDF.Container = Frame1 'Attach the PDF Browser control to a container.

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub



Private Sub Form_Resize()
On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500
        
        b8TitleBar1.Width = Me.ScaleWidth
        Frame1.Width = Me.ScaleWidth
        Frame1.Height = Me.ScaleHeight - 1000
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
    MAIN.RemoveChild Me.Name
    'The Browser control will load a blank
    m_objPDF.LoadFile ""
    
    'Set object to nothing
    Set m_objPDF = Nothing
    Set frmProduction = Nothing
End Sub

Private Sub Form_Activate()
On Error Resume Next
With MAIN
    Me.BackColor = .ACPMenu.BackColor
    Me.Picture = .ACPMenu.LoadBackground
    Frame1.BackColor = .ACPMenu.BackColor
End With
'LoadPDF App.Path & "\Picture\CCP\" & p_int_part_1 & ".pdf", izoomCCP
LoadPDF App.Path & "\User Manual.pdf", izoomWI
MAIN.ActivateChild Me

End Sub

