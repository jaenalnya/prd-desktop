VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptSubSPB2 
   Caption         =   "Sub Report Idle Time"
   ClientHeight    =   12375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   40217
   _ExtentY        =   21828
   SectionData     =   "RptSubTotalProd.dsx":0000
End
Attribute VB_Name = "RptSubSPB2"
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

            txtperiode.Text = Format(.Fields("period_shift").Value, "yyyy-mm-dd")
            txtshift.Text = .Fields("shift").Value

            txtok.Text = .Fields("ok").Value
            txtsisa.Text = .Fields("sisa").Value
            txthold.Text = .Fields("hold").Value
            txttotal.Text = .Fields("sub_total").Value
        End If
    End With

End Sub

