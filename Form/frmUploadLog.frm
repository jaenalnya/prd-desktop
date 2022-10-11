VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmUploadLog 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7935
   Icon            =   "frmUploadLog.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   1440
      TabIndex        =   14
      Top             =   2025
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   9043971
      CurrentDate     =   43347
   End
   Begin VB.ComboBox cboPlant 
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
      Left            =   1440
      TabIndex        =   12
      Top             =   1035
      Width           =   3930
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   90
      Top             =   3060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   180
      TabIndex        =   10
      Top             =   3690
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin lvButton.lvButtons_H cmdPath 
      Height          =   375
      Left            =   7155
      TabIndex        =   9
      Top             =   1530
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   661
      Caption         =   "..."
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
   Begin VB.TextBox Txtpath 
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
      Left            =   1440
      TabIndex        =   8
      Text            =   "Select Path of Database"
      Top             =   1530
      Width           =   5640
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   7935
      TabIndex        =   2
      Top             =   0
      Width           =   7935
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "UPLOAD LOG MESIN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   945
         TabIndex        =   4
         Top             =   135
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Menu ini digunakan untuk import data log  mesin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   960
         TabIndex        =   3
         Top             =   405
         Width           =   3570
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   45
         Picture         =   "frmUploadLog.frx":169B2
         Top             =   45
         Width           =   720
      End
   End
   Begin PRD.Liner Liner2 
      Height          =   30
      Left            =   75
      TabIndex        =   0
      Top             =   2910
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   53
   End
   Begin PRD.Liner Liner1 
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   780
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   53
   End
   Begin lvButton.lvButtons_H cmdSave 
      Height          =   435
      Left            =   4635
      TabIndex        =   5
      Top             =   3060
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   767
      Caption         =   "&Upload [F5]"
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
      Image           =   "frmUploadLog.frx":2D364
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   435
      Left            =   6255
      TabIndex        =   6
      Top             =   3060
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   767
      Caption         =   "&Close [ESC]"
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
      Image           =   "frmUploadLog.frx":334EE
      cBack           =   -2147483633
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   330
      Left            =   3690
      TabIndex        =   15
      Top             =   2025
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   9043971
      CurrentDate     =   43347
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3285
      TabIndex        =   16
      Top             =   2070
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   13
      Top             =   2070
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Location Plant "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   11
      Top             =   1035
      Width           =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Location Files"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   7
      Top             =   1575
      Width           =   1320
   End
End
Attribute VB_Name = "frmUploadLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim CNewDB                         As ADODB.Connection





Private Sub cboPlant_KeyPress(KeyAscii As Integer)
    KeyAscii = AutoMatchCBBox(cboPlant, KeyAscii)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPath_Click()
'On Error GoTo errhndl

  With Me.CommonDialog1
    .DialogTitle = "Select Database Access"
    .InitDir = App.Path
    .Filter = "Database Access (*.mdb)|*.mdb|All Files (*.*)|*.*"
    .ShowOpen
    Me.Txtpath.Text = .FileName
  End With
   
  If Txtpath.Text <> "" Then
    
    Set CNewDB = New ADODB.Connection
    With CNewDB
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Txtpath.Text & "; Persist Security Info=False;Jet OLEDB:Database Password=eagle@#$%02"
        .Open
        .CursorLocation = adUseClient
    End With
    
    cmdSave.Enabled = True
    
  Else
    Txtpath.Text = "Select Path of Database"
    cmdSave.Enabled = True
  End If
  
  Exit Sub

errhndl:
  MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdSave_Click()
On Error GoTo errhndl

  Dim datMDB    As New ADODB.Recordset
  Dim sSQL      As String
  Dim sMDB      As String
  Dim j         As Double
  
  sMDB = "SELECT * From etcom Where empl_code <> '0009' And acc_code <> 'y' " & _
              "AND tr_date BETWEEN #" & Format(DTPicker1.Value, "m/d/yyyy") & "# " & _
              "AND #" & Format(DTPicker2.Value, "m/d/yyyy") & "# ORDER BY empl_code,tr_date,acc_code"
              
  datMDB.Open sMDB, CNewDB, adOpenDynamic, adLockOptimistic
  
  With datMDB
    If Not (.BOF And .EOF) Then
        Screen.MousePointer = vbHourglass
        .MoveFirst
         While Not .EOF
         j = j + 1
            
             sSQL = "INSERT INTO habsen (tr_no,loc_code,remoteno,tr_date,tr_time,empl_code,acc_code) VALUES ( " & _
                    "" & j & ",'" & datMDB!loc_code & "','" & datMDB!remoteno & "','" & Format(datMDB!tr_date, "yyyy-mm-dd") & "', " & _
                    "'" & Format(datMDB!tr_date, "hh:mm:ss") & "','" & datMDB!empl_code & "','" & datMDB!acc_code & "')"
             
             sSQL_Insert sSQL
             
'             CN.Execute "INSERT INTO hrd_mesin_results (sys_plant_id,pin, tanggal, jam, " & _
'                    "mode,dateencoded,encodedby) VALUES ( " & _
'                    Mid(cboPlant, 1, 1) & ",'" & datMDB!empl_code & "','" & _
'                    Format(datMDB!tr_date, "yyyy-mm-dd hh:mm:ss") & "','" & Format(datMDB!tr_date, "hh:mm") & "'," & _
'                    datMDB!acc_code & ",'" & _
'                    Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER.USERNAME & "')"

            .MoveNext
             ProgressBar1.Value = j / datMDB.RecordCount * 100
          Wend
    Else
        MsgBox "Data is empty", vbInformation, "Data Empty"
        datMDB.Close
      
        cmdSave.Enabled = True

        Exit Sub
    End If

  End With
  
  MsgBox "Copy data finished", vbExclamation

  
  Screen.MousePointer = vbDefault

  
  datMDB.Close
  Set datMDB = Nothing
  
  CNewDB.Close
  Set CNewDB = Nothing
  
  cmdSave.Enabled = False
  Exit Sub

errhndl:
  Resume Next
End Sub

Private Sub Form_Load()
    'AddComboField "sys_plants", "id", "name", cboPlant
    DTPicker1.Value = Format(Now, "dd-MMM-yyyy")
    DTPicker2.Value = Format(Now, "dd-MMM-yyyy")
    
End Sub
