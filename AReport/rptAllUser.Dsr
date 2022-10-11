VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptAllUser 
   Caption         =   "User Master File"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "rptAllUser.dsx":0000
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19315
   SectionData     =   "rptAllUser.dsx":054A
End
Attribute VB_Name = "rptAllUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Detail_Format()
On Error Resume Next
    With DTRpt.Recordset
        If Not .EOF Then
            txtNo.Text = Val(txtNo.Text) + 1
            txtNo.Text = txtNo.Text & "."
        End If
    End With
End Sub

Private Sub GroupHeader1_Format()
On Error Resume Next
    With DTRpt.Recordset
        If Not .EOF Then
            txtNo.Text = "0"
        End If
    End With
End Sub
