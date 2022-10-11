VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LvButton.ocx"
Begin VB.Form frmChartMonitoring 
   Caption         =   "Chart Monitoring"
   ClientHeight    =   10140
   ClientLeft      =   5070
   ClientTop       =   2145
   ClientWidth     =   17130
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10140
   ScaleWidth      =   17130
   Begin VB.PictureBox picFooter 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   17130
      TabIndex        =   0
      Top             =   9735
      Width           =   17130
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   5175
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   1
         Top             =   45
         Width           =   4150
         Begin VB.CommandButton btnFirst 
            Height          =   315
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "First 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnPrev 
            Height          =   315
            Left            =   3075
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Previous 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnLast 
            Height          =   315
            Left            =   3705
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Last 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnNext 
            Height          =   315
            Left            =   3390
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Next 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.Label lblPageInfo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 0 of 0"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   60
            Width           =   2535
         End
      End
      Begin PRD.Liner Liner1 
         Height          =   30
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   53
      End
      Begin VB.Label lblRecSum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Records"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   690
      End
   End
   Begin lvButton.lvButtons_H cmdClose 
      Height          =   405
      Left            =   90
      TabIndex        =   9
      Top             =   585
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   714
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
      cFHover         =   4210752
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Image           =   "FrmChartMonitoring.frx":0000
      cBack           =   16119285
   End
   Begin MSComCtl2.DTPicker DTAwal 
      Height          =   330
      Left            =   2250
      TabIndex        =   10
      Top             =   45
      Width           =   1590
      _ExtentX        =   2805
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
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   87359491
      CurrentDate     =   43642
   End
   Begin MSComCtl2.DTPicker DTAkhir 
      Height          =   330
      Left            =   4410
      TabIndex        =   11
      Top             =   45
      Width           =   1590
      _ExtentX        =   2805
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
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   87359491
      CurrentDate     =   43642
   End
   Begin PRD.ucChartBar ucChartBar1 
      Height          =   8520
      Left            =   1575
      TabIndex        =   14
      Top             =   585
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   15028
      Title           =   "Monitoring Leader"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1620
      TabIndex        =   13
      Top             =   90
      Width           =   600
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4005
      TabIndex        =   12
      Top             =   90
      Width           =   510
   End
End
Attribute VB_Name = "frmChartMonitoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim srcItem                         As ListItem
Dim srcRecord                       As String
Dim srcMonitoring                           As Variant
Dim srcSQL                          As String
Dim RecordPage                      As New clsPaging
Dim SQLParser                       As New clsSQLSelectParser

Private Sub DTAwal_Click()

    Call ChartData
End Sub

Private Sub DTAwal_CloseUp()
    Call ChartData
End Sub

Private Sub cmdClose_Click()
    CommandPass "Close"
End Sub


Private Sub cmdKeluar_Click()
    CommandPass "Keluar"
End Sub

Private Sub cmdMasuk_Click()
    CommandPass "Masuk"
End Sub

Private Sub cmdRefresh_Click()
    CommandPass "Refresh"
End Sub

Private Sub CmdSearch_Click()
    CommandPass "Search"
End Sub

Private Sub Form_Activate()
On Error Resume Next
With MAIN
    Me.BackColor = .ACPMenu.BackColor
    Me.Picture = .ACPMenu.LoadBackground
    picFooter.BackColor = .ACPMenu.BackColor
    Picture2.BackColor = .ACPMenu.BackColor
End With

'AltLVBackground lvList, vbWhite, &H8000000F, frmMonitoringLeader
MAIN.ActivateChild Me


End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler

DTAwal.Value = Format(Now, "dd/MMM/yyyy")
DTAkhir.Value = Format(Now, "dd/MMM/yyyy")

Call ChartData

Exit Sub
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation
End Sub

Private Sub ChartData()
On Error Resume Next
    Dim Value As Collection
    Dim i As Long, j As Long
    Dim Palette() As String
    Dim Users() As String
    Dim MyArray() As String
    Dim Rs As New ADODB.Recordset
    Dim sSQL As Variant

    sSQL = "select a.hrd_employee_id,b.name as leader_name, count(a.prod_machine_id) jumlah_check"
    sSQL = sSQL & " from prod_monitoring_leaders a"
    sSQL = sSQL & " left join hrd_employees b on a.hrd_employee_id = b.id"
    sSQL = sSQL & " where a.period_shift between '" & Format(DTAwal.Value, "yyyy-mm-dd") & "' and '" & Format(DTAkhir.Value, "yyyy-mm-dd") & "' and a.plant_mark = '" & p_plant_mark & "'"
    sSQL = sSQL & " group by a.hrd_employee_id"
    sSQL = sSQL & " order by b.name"

    Rs.Open sSQL, CN, adOpenStatic, adLockOptimistic
    
    ucChartBar1.LabelsVisible = True
    ucChartBar1.Clear

    i = 0
    If Rs.RecordCount > 0 Then
        Do While Not Rs.EOF
            ReDim Preserve Users(i)
            ReDim Preserve MyArray(i)
            
            Users(i) = Rs.Fields("leader_name")
            MyArray(i) = Rs.Fields("jumlah_check")
            
            Rs.MoveNext
            i = i + 1
        Loop
    
        Palette = Split("&HFF8D11,&HA744E0,&H376CE6,&H40AB1A,&H7B006B,&H7B006B,&H7B006B", ",")
    
        Set Value = New Collection
        For i = 0 To UBound(Users)
            Value.Add Users(i)
        Next
        
        ucChartBar1.AddAxisItems Value, False, 0, ccEnter
    
    
        Set Value = New Collection
        For j = 0 To UBound(MyArray)
            Value.Add MyArray(j)
        Next
    
        ucChartBar1.AddSerie "Leader", CLng(Palette(0)), Value
        ucChartBar1.Refresh
    End If
    
End Sub
Private Sub Form_Resize()
On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500

        Liner1.Width = ScaleWidth
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    MAIN.RemoveChild Me.Name
    Set frmChartMonitoring = Nothing
End Sub


Public Sub CommandPass(ByVal srcPerformWhat As String)
On Error GoTo errPerformWhat
Select Case srcPerformWhat

    Case "Close" 'Close
            Unload Me
End Select
Exit Sub
errPerformWhat:
     MsgBox "Error Number:" & Err.Number & vbNewLine & _
            "Description:" & Err.Description, vbExclamation
End Sub


