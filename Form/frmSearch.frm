VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pencarian Data"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   5625
      TabIndex        =   9
      Top             =   1710
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   661
      Caption         =   "E&xit"
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
      Image           =   "frmSearch.frx":0A02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdOk 
      Height          =   375
      Left            =   4230
      TabIndex        =   8
      Top             =   1710
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   661
      Caption         =   "Cari"
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
      Image           =   "frmSearch.frx":6B8C
      cBack           =   -2147483633
   End
   Begin VB.ComboBox cmbFields 
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
      ItemData        =   "frmSearch.frx":CD16
      Left            =   1800
      List            =   "frmSearch.frx":CD18
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4995
   End
   Begin VB.Frame Frame1 
      Caption         =   " Condition "
      Height          =   975
      Left            =   90
      TabIndex        =   1
      Top             =   600
      Width           =   6705
      Begin VB.TextBox txtFilter 
         Height          =   330
         Index           =   0
         Left            =   3105
         TabIndex        =   3
         Top             =   360
         Width           =   3390
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   330
         Index           =   0
         Left            =   3120
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   1470
         _ExtentX        =   2593
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
         CustomFormat    =   "MMM-dd-yyyy"
         Format          =   120913923
         CurrentDate     =   38207
      End
      Begin VB.ComboBox cmbOperation 
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
         ItemData        =   "frmSearch.frx":CD1A
         Left            =   240
         List            =   "frmSearch.frx":CD36
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2470
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   330
         Index           =   1
         Left            =   5040
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   1470
         _ExtentX        =   2593
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
         CustomFormat    =   "MMM-dd-yyyy"
         Format          =   120913923
         CurrentDate     =   38207
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "And"
         Height          =   255
         Left            =   4605
         TabIndex        =   7
         Top             =   390
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   2760
         Stretch         =   -1  'True
         Top             =   405
         Width           =   240
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pilih Pencarian Data"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   165
      Width           =   1710
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Public srcColumnHeaders As ColumnHeaders 'Source column headers
Public srcNoOfCol As Long
Public srcForm As Form 'Source form

Private Sub cmbOperation_Click(Index As Integer)
    If Index = 0 Then
        If cmbOperation(Index).ListIndex = 7 Then
            dtpDate(0).Visible = True
            dtpDate(1).Visible = True
            txtFilter(0).Visible = False
        Else
            txtFilter(0).Visible = True
            dtpDate(0).Visible = False
            dtpDate(1).Visible = False
        End If
    Else
        If cmbOperation(Index).ListIndex = 7 Then
            dtpDate(2).Visible = True
            dtpDate(3).Visible = True
            txtFilter(1).Visible = False
        Else
            txtFilter(1).Visible = True
            dtpDate(2).Visible = False
            dtpDate(3).Visible = False
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    'Verify
    If cmbOperation(0).ListIndex <> 7 Then If txtFilter(0).Text = "" Then txtFilter(0).SetFocus: Exit Sub
    
    On Error GoTo Err
    Dim strFilter As String
    'Initialize the fields
    strFilter = Replace(cmbFields.Text, "/", "") 'ex. City/Town for tblCustomer
    strFilter = Replace(cmbFields.Text, " ", "")
    strFilter = "" & strFilter & ""
    'Initialize the operation used
    'First operation
    Select Case cmbOperation(0).ListIndex
        Case 0: strFilter = strFilter & " LIKE '%" & txtFilter(0).Text & "%'"
        Case 1: strFilter = strFilter & " = '" & txtFilter(0).Text & "'"
        Case 2: strFilter = strFilter & " <> '" & txtFilter(0).Text & "'"
        Case 3: strFilter = strFilter & " > '" & txtFilter(0).Text & "'"
        Case 4: strFilter = strFilter & " >= '" & txtFilter(0).Text & "'"
        Case 5: strFilter = strFilter & " < '" & txtFilter(0).Text & "'"
        Case 6: strFilter = strFilter & " <= '" & txtFilter(0).Text & "'"
        Case 7: strFilter = strFilter & " BETWEEN '" & Format(dtpDate(0).Value, "yyyy-mm-dd") & "' AND '" & Format(dtpDate(1).Value, "yyyy-mm-dd") & "'"
    End Select

        
    'InputBox "", , strFilter
    'Pass the condition to filtered records
    srcForm.FilterRecord strFilter
    'Clear used variables
    strFilter = vbNullString
    
    Unload Me
    Exit Sub
Err:
        If Err.Number = -2147352571 Then
            MsgBox "Invalid search operation.", vbExclamation
            Unload Me
        ElseIf Err.Number = 3001 Then
            Resume Next
        Else
            Prompt_Err Err, "frmFilter", "cmdOk_Click"
        End If
End Sub

Private Sub Form_Load()
    'Initialize values
    dtpDate(0).Value = Date
    dtpDate(1).Value = Date

    'Set the images for the controls
    With MAIN
        Image1.Picture = .i16x16.ListImages(7).Picture
    End With
    
    Dim i As Integer
    If srcNoOfCol = 0 Then srcNoOfCol = srcColumnHeaders.Count
    
    For i = 1 To srcNoOfCol
        If srcColumnHeaders(i).Text <> "" Then cmbFields.AddItem srcColumnHeaders(i).Text
    Next i
    i = 0
    
    cmbFields.ListIndex = 0
    cmbOperation(0).ListIndex = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSearch = Nothing
End Sub

Private Sub txtFilter_GotFocus(Index As Integer)
    HLText txtFilter(Index)
End Sub
