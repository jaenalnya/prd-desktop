VERSION 5.00
Begin VB.Form frmMAIN 
   Caption         =   "TSSI Auto Update"
   ClientHeight    =   3510
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   6615
   Icon            =   "FrmMAIN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3510
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu smUpload 
         Caption         =   "Upload"
      End
      Begin VB.Menu smDownload 
         Caption         =   "Download"
      End
      Begin VB.Menu smGaris 
         Caption         =   "-"
      End
      Begin VB.Menu smExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************
'*                  THANKS FOR VISIT http://dlaboratory.wordpress.com               *
'*                          by Admin of dRecks Laboratory                           *
'*                      Email   : dRecks.Lab@gmail.com                              *
'*                      Facebook: http://facebook.com/dRecks.Lab                    *
'************************************************************************************

Option Explicit
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Enum TrayRetunEventEnum
    MouseMove = &H200
    LeftUp = &H202
    LeftDown = &H201
    LeftDblClick = &H203
    RightUp = &H205
    RightDown = &H204
    RightDblClick = &H206
    MiddleUp = &H208
    MiddleDown = &H207
    MiddleDblClick = &H209
    BalloonClick = &H405
    BalloonClose = &H404
End Enum

Dim Tray As NOTIFYICONDATA

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TrayEvent As TrayRetunEventEnum
TrayEvent = X / Screen.TwipsPerPixelX
        
Select Case TrayEvent
    Case MouseMove
    Case LeftUp
        Me.Show
        Me.WindowState = 0
    Case LeftDown
    Case LeftDblClick
    Case MiddleUp
    Case MiddleDown
    Case MiddleDblClick
    Case RightUp
    Case RightDown
        PopupMenu Menu
    Case RightDblClick
    Case BalloonClick
    Case BalloonClose
End Select
End Sub

Private Sub Form_Resize()
    If WindowState = 1 Then Minimize_To_Tray
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, Tray
End Sub

Private Sub Menu1_Click()
    MsgBox "Anda Klik Menu 1"
End Sub

Private Sub Menu2_Click()
    MsgBox "Anda Klik Menu 2"
End Sub

Private Sub Minimize_To_Tray()
    Tray.cbSize = Len(Tray)
    Tray.hWnd = Me.hWnd
    Tray.uId = 1&
    Tray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    Tray.ucallbackMessage = WM_MOUSEMOVE
    Tray.hIcon = Me.Icon
    Tray.szTip = "Membuat Tray Icon" & Chr$(0)
    Shell_NotifyIcon NIM_ADD, Tray
    Me.Hide
End Sub
