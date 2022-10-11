VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptAllSupplier 
   Caption         =   "Supplier Masterlist"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "rptAllSupplier.dsx":0000
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19288
   SectionData     =   "rptAllSupplier.dsx":054A
End
Attribute VB_Name = "rptAllSupplier"
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

