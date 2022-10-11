VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptOperator 
   Caption         =   "Laporan Operator"
   ClientHeight    =   12375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   40217
   _ExtentY        =   21828
   SectionData     =   "RptOperator.dsx":0000
End
Attribute VB_Name = "RptOperator"
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
 
            txtNikOpt1.Text = .Fields("nik_operator_1").Value
            txtNikOpt2.Text = .Fields("nik_operator_2").Value
            txtNamaOpt1.Text = .Fields("name_operator_1").Value
            txtNamaOpt2.Text = .Fields("name_operator_2").Value

        End If
    End With

End Sub

