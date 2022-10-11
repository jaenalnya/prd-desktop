VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form FrmTrial 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4920
   Icon            =   "FrmTrial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtGen 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   9
      Top             =   2235
      Width           =   4770
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1635
      Width           =   4770
   End
   Begin lvButton.lvButtons_H CmdAbout 
      Height          =   375
      Left            =   1755
      TabIndex        =   7
      Top             =   3015
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      Caption         =   "&About"
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
      Image           =   "FrmTrial.frx":6852
      cBack           =   -2147483633
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   4920
      TabIndex        =   0
      Top             =   0
      Width           =   4920
      Begin VB.PictureBox Liner1 
         Height          =   30
         Left            =   0
         ScaleHeight     =   30
         ScaleWidth      =   10215
         TabIndex        =   1
         Top             =   960
         Width           =   10215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "R A F I S M A   S Y S T E M ®™"
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
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   3015
      End
      Begin VB.Image Image1 
         Height          =   780
         Left            =   120
         Picture         =   "FrmTrial.frx":69E1
         Top             =   120
         Width           =   720
      End
   End
   Begin lvButton.lvButtons_H CmdEntSerial 
      Height          =   390
      Left            =   225
      TabIndex        =   3
      Top             =   3015
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   688
      Caption         =   "&Enter Serial"
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
      Image           =   "FrmTrial.frx":77CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   390
      Left            =   3330
      TabIndex        =   4
      Top             =   3015
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   688
      Caption         =   "&Exit"
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
      Image           =   "FrmTrial.frx":E030
      cBack           =   -2147483633
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   45
      TabIndex        =   12
      Top             =   3645
      Width           =   375
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Code:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   90
      TabIndex        =   11
      Top             =   1995
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Serial Number :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   90
      TabIndex        =   10
      Top             =   1395
      Width           =   1845
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © By Rafisma.com. All Rights Reserved 2016"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   495
      TabIndex        =   6
      Top             =   3555
      Width           =   3990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTRATION SOFTWARE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Top             =   990
      Width           =   3255
   End
End
Attribute VB_Name = "FrmTrial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Dim dwLen                            As Long
Dim strString                        As String
Dim clsDS2                           As New clsDS2

Private Sub CmdAbout_Click()

    'Show details about your software.
        MsgBox "Company Name: " & App.CompanyName & vbCrLf & "Product Name: " & App.ProductName & vbCrLf & "Version: " & App.Major & "." & App.Revision & "." & App.Minor & vbCrLf & vbCrLf & "Little message about your product here.."

End Sub


Private Sub cmdExit_Click()
    'Terminate the program if the user decides to.
        Unload Me
    End

End Sub


Private Sub Form_Load()

    On Error Resume Next
    'Create a buffer
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    'Get the computer name

    GetComputerName strString, dwLen
    strString = Left(strString, dwLen)
    txtUsername.Text = KeyGen(strString, "pramasystem", 3)

    Dim Line01 As String
    Dim Line02 As String
    
    'Open trial config file to check if the software is registered or not.
    Open App.Path & "\WinJT.001" For Input As #1
    
    
    'Grab details from config file.
    Line Input #1, Line01
    Line Input #1, Line02
    Close #1
    
    'Decrypt the text using DS2 Cipher decryption.
    Line01 = clsDS2.DecryptString(Line01, "589501068402658", True)
    Line02 = clsDS2.DecryptString(Line02, "589501068402658", True)
    
    'Check to see if the text matches a valid registration code.
    If KeyGen(Line01, "pramasystem", 3) = Line02 Then

    Dim rc As Long

    If App.PrevInstance Then
        rc = MsgBox("Application is already running", vbCritical, App.Title)
        Unload Me
    Else
        MAIN.Show
        Unload Me
    End If

    End If
    
    
End Sub

Private Sub CmdEntSerial_Click()

    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    'Get the computer name

    GetComputerName strString, dwLen
    strString = Left(strString, dwLen)
    
'Check to see if the user input matches correct information.
    If KeyGen(txtUsername, "pramasystem", 3) = txtGen Then
    
        'EmailSend strString
        
        'Encrypt the file to stop people from looking at this hidden info.
            txtUsername.Text = clsDS2.EncryptString(txtUsername.Text, "589501068402658", True)
            txtGen.Text = clsDS2.EncryptString(txtGen.Text, "589501068402658", True)
        
        'Write the details to file, if they are correct then the software will be registered.
        Open App.Path & "\WinJT.001" For Output As #1
            Print #1, txtUsername.Text
            Print #1, txtGen.Text
        Close
        
        MsgBox "Registration successfull, thank you for purchasing this product, you will need to re-launch this program for the changes to take effect.", vbInformation, "Registration Complete!"
        Unload Me
    End If
End Sub



Private Sub EmailSend(sFile As String)
On Error Resume Next
    Dim retVal          As String
    
    eTo = "jaenalnya@gmail.com"
    eCC = "jaenal_thea@yahoo.co.id"
    eFromName = "<PRAMA SYSTEM>"
    eFromEmail = "jaenalnya@gmail.com"
    eServer = "smtp.gmail.com"
    ePort = 465
    eUsername = "jaenalnya@gmail.com"
    ePassword = "jaenal25031985"
    eSSL = 1

    retVal = SendMail(eTo, _
        "Informasi Pramasystem", _
        eFromName & eFromEmail, _
        "Dear,<br><br>Ini adalah email dari software Prama System<br>" & _
        "Software Prama mengirimkan file<br> Nama File : " & sFile & "<br><br>Thanks,<br>PramaSystem<br>do not reply to this message", _
        eServer, _
        ePort, _
        eUsername, _
        ePassword, _
        eSSL, eCC)
        Label3.Caption = IIf(retVal = "ok", "OK!", retVal)
End Sub
