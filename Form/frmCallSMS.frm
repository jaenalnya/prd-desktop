VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmCallSMS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panggilan"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9975
   ControlBox      =   0   'False
   Icon            =   "frmCallSMS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2805
      Left            =   4995
      TabIndex        =   8
      Top             =   0
      Width           =   4920
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   4275
         Top             =   405
      End
      Begin lvButton.lvButtons_H cmdQC 
         Height          =   555
         Left            =   1260
         TabIndex        =   9
         Top             =   2070
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   979
         Caption         =   "START CALL"
         CapAlign        =   2
         BackStyle       =   1
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
         cBack           =   8421376
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PANGGIL QC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   420
         Left            =   945
         TabIndex        =   16
         Top             =   180
         Width           =   2715
      End
      Begin VB.Label lblLossTime 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   1
         Left            =   1980
         TabIndex        =   14
         Top             =   1620
         Width           =   2715
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LOSS TIME :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   135
         TabIndex        =   13
         Top             =   1620
         Width           =   1770
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TIME STOP :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   90
         TabIndex        =   12
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblTimeStart 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Index           =   1
         Left            =   1980
         TabIndex        =   11
         Top             =   720
         Width           =   2715
      End
      Begin VB.Label lblTimeStop 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   1
         Left            =   1980
         TabIndex        =   10
         Top             =   1170
         Width           =   2715
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2805
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   4920
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   4275
         Top             =   2025
      End
      Begin lvButton.lvButtons_H cmdTeknisi 
         Height          =   555
         Left            =   1260
         TabIndex        =   7
         Top             =   2070
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   979
         Caption         =   "START CALL"
         CapAlign        =   2
         BackStyle       =   1
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
         cBack           =   16576
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PANGGIL TEKNISI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   420
         Left            =   990
         TabIndex        =   15
         Top             =   225
         Width           =   2715
      End
      Begin VB.Label lblTimeStop 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   0
         Left            =   1980
         TabIndex        =   6
         Top             =   1170
         Width           =   2715
      End
      Begin VB.Label lblTimeStart 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Index           =   0
         Left            =   1980
         TabIndex        =   5
         Top             =   720
         Width           =   2715
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "TIME STOP :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LOSS TIME :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   135
         TabIndex        =   3
         Top             =   1620
         Width           =   1770
      End
      Begin VB.Label lblLossTime 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   0
         Left            =   1980
         TabIndex        =   2
         Top             =   1620
         Width           =   2715
      End
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   555
      Left            =   4140
      TabIndex        =   0
      Top             =   3015
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   979
      Caption         =   "EXIT"
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
      Image           =   "frmCallSMS.frx":617A
      cBack           =   4210752
   End
End
Attribute VB_Name = "frmCallSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Dim eng_prod_1 As String
Dim eng_prod_2 As String
Dim sQL As String
Dim sYourCommand As String
Dim Relay As String


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdQC_Click()
On Error Resume Next
    If cmdQC.Caption = "START CALL" Then
        Timer3.Enabled = True
        lblTimeStart(1).Caption = Format(Now, "yyyy-mm-dd hh:mm:ss")
        cmdQC.Caption = "STOP CALL"

        'insert data
        sQL = "insert into sip_production.prod_call_logs"
        sQL = sQL & " (plant_mark,prod_machine_id,mkt_customer_id,eng_product_1,eng_product_2,"
        sQL = sQL & " period_shift,start_call,created_at,created_by,description) values"
        sQL = sQL & " ('" & p_plant_mark & "','" & p_prod_machine_id & "','" & p_mkt_customer_id & "'," & eng_prod_1 & ""
        sQL = sQL & " ," & eng_prod_2 & ",'" & Format(p_shift, "yyyy-mm-dd") & "','" & lblTimeStart(1).Caption & "'"
        sQL = sQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER.KODEUSER & "','CALL QC')"
        
        sSQL_Insert sQL
        
        cmdExit.Enabled = False
            
        sYourCommand = "CommandApp_USBRelay  " & Relay & " open 02"
        Shell "C:\Windows\System32\cmd.exe /c" & sYourCommand, vbHide
        
    Else
        Timer3.Enabled = False
        cmdQC.Caption = "START CALL"


        sQL = "update sip_production.prod_call_logs set end_call = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "'"
            sQL = sQL & " ,idle_time = '" & lblLossTime(1).Caption & "' where "
            sQL = sQL & " plant_mark = '" & p_plant_mark & "' "
            sQL = sQL & " and prod_machine_id = '" & p_prod_machine_id & "'"
            sQL = sQL & " and mkt_customer_id = '" & p_mkt_customer_id & "'"
            sQL = sQL & " and start_call = '" & lblTimeStart(1).Caption & "' and description = 'CALL QC'"
            
        sSQL_Update sQL
        
        cmdExit.Enabled = True
                
        sYourCommand = "CommandApp_USBRelay  " & Relay & " close 02"
        Shell "C:\Windows\System32\cmd.exe /c" & sYourCommand, vbHide
        
    End If
End Sub


Private Sub cmdTeknisi_Click()
On Error Resume Next
'CommandApp_USBRelay  afEd5 open 01
'CommandApp_USBRelay  afEd5 close 01
'CommandApp_USBRelay  afEd5 open 255

    If cmdTeknisi.Caption = "START CALL" Then
        Timer1.Enabled = True
        lblTimeStart(0).Caption = Format(Now, "yyyy-mm-dd hh:mm:ss")
        cmdTeknisi.Caption = "STOP CALL"

        'insert data
        sQL = "insert into sip_production.prod_call_logs"
        sQL = sQL & " (plant_mark,prod_machine_id,mkt_customer_id,eng_product_1,eng_product_2,"
        sQL = sQL & " period_shift,start_call,created_at,created_by,description) values"
        sQL = sQL & " ('" & p_plant_mark & "','" & p_prod_machine_id & "','" & p_mkt_customer_id & "'," & eng_prod_1 & ""
        sQL = sQL & " ," & eng_prod_2 & ",'" & Format(p_shift, "yyyy-mm-dd") & "','" & lblTimeStart(0).Caption & "'"
        sQL = sQL & " ,'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & ACTIVE_USER.KODEUSER & "','CALL TEKNISI')"
        
        sSQL_Insert sQL
        
        cmdExit.Enabled = False
        
        sYourCommand = "CommandApp_USBRelay  " & Relay & " open 01"
        Shell "C:\Windows\System32\cmd.exe /c" & sYourCommand, vbHide
        
    Else
        Timer1.Enabled = False
        cmdTeknisi.Caption = "START CALL"


        sQL = "update sip_production.prod_call_logs set end_call = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "'"
            sQL = sQL & " ,idle_time = '" & lblLossTime(0).Caption & "' where "
            sQL = sQL & " plant_mark = '" & p_plant_mark & "' "
            sQL = sQL & " and prod_machine_id = '" & p_prod_machine_id & "'"
            sQL = sQL & " and mkt_customer_id = '" & p_mkt_customer_id & "'"
            sQL = sQL & " and start_call = '" & lblTimeStart(0).Caption & "' and description = 'CALL TEKNISI'"
            
        sSQL_Update sQL
        
        cmdExit.Enabled = True
                
        sYourCommand = "CommandApp_USBRelay  " & Relay & " close 01"
        Shell "C:\Windows\System32\cmd.exe /c" & sYourCommand, vbHide
        
    End If
    
End Sub


Private Sub Form_Activate()
On Error Resume Next
With MAIN
    Me.BackColor = .ACPMenu.BackColor
    Me.Picture = .ACPMenu.LoadBackground

End With

End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    If p_eng_product_1 = "" Then
        eng_prod_1 = "NULL"
    Else
        eng_prod_1 = p_eng_product_1
    End If
    If p_eng_product_2 = "" Then
        eng_prod_2 = "NULL"
    Else
        eng_prod_2 = p_eng_product_2
    End If

    Relay = ReadINI("SETTING", "RELAY", App.Path & "\Settings.ini")
            
Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
    
End Sub

Private Sub Timer1_Timer()

    lblTimeStop(0).Caption = Format(Now, "yyyy-mm-dd hh:mm:ss")
    lblLossTime(0).Caption = Format(CDate(CDate(lblTimeStop(0).Caption) - CDate(lblTimeStart(0).Caption)), "hh:mm:ss")

End Sub


Private Sub Timer3_Timer()

    lblTimeStop(1).Caption = Format(Now, "yyyy-mm-dd hh:mm:ss")
    lblLossTime(1).Caption = Format(CDate(CDate(lblTimeStop(1).Caption) - CDate(lblTimeStart(1).Caption)), "hh:mm:ss")

End Sub

