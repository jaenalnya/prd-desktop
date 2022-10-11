VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptSubIdle 
   Caption         =   "Sub Report Idle Time"
   ClientHeight    =   12375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   40217
   _ExtentY        =   21828
   SectionData     =   "RptSubIdle.dsx":0000
End
Attribute VB_Name = "RptSubIdle"
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
 
            txtStop.Text = Format(.Fields("start_idle").Value, "hh:mm:ss")
            txtJalan.Text = Format(.Fields("end_idle").Value, "hh:mm:ss")
            txtidle.Text = Format(.Fields("idle_time").Value, "hh:mm:ss")
            txtKode.Text = .Fields("kode").Value
            txtKeterangan.Text = .Fields("idle_name").Value
            txtkaryawan.Text = .Fields("description").Value

        End If
    End With

End Sub

