VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activation DwS"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   510
   ClientWidth     =   5040
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0EE2
   ScaleHeight     =   1350
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4050
      TabIndex        =   1
      Top             =   945
      Width           =   855
   End
   Begin VB.CommandButton cmdKGen 
      Caption         =   "Key Generator"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   990
      TabIndex        =   0
      Top             =   120
      Width           =   3195
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim clsDS2 As New clsDS2

Private Sub CmdAbout_Click()
    'Show details about your software.
        MsgBox "Company Name: " & App.CompanyName & vbCrLf & "Product Name: " & App.ProductName & vbCrLf & "Version: " & App.Major & "." & App.Revision & "." & App.Minor & vbCrLf & vbCrLf & "Little message about your product here.."
End Sub


Private Sub cmdExit_Click()
    'Terminate the program if the user decides to.
        End
End Sub

Private Sub CmdKGen_Click()
    'Load the Key Generator form.
        frmKeyGen.Show
End Sub


