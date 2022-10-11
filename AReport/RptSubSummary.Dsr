VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptSubSummary 
   Caption         =   "Sub Report Idle Time"
   ClientHeight    =   12375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   40217
   _ExtentY        =   21828
   SectionData     =   "RptSubSummary.dsx":0000
End
Attribute VB_Name = "RptSubSummary"
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

            txtcavity.Text = .Fields("cavity").Value
            txtshot.Text = .Fields("total").Value
            txtgross.Text = .Fields("gross").Value
            txtng.Text = .Fields("total_ng").Value
            txtnet.Text = .Fields("net_produksi").Value
            txtProdYield.Text = .Fields("prod_yield").Value
        End If
    End With

End Sub

