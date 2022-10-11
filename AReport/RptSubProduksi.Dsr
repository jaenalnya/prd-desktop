VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptSubProduksi 
   Caption         =   "Laporan Product"
   ClientHeight    =   12375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   40217
   _ExtentY        =   21828
   SectionData     =   "RptSubProduksi.dsx":0000
End
Attribute VB_Name = "RptSubProduksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Detail_Format()
On Error Resume Next
    With DTRpt.Recordset
        If Not .EOF Then
            'txtNo.Text = Val(txtNo.Text) + 1
            'txtNo.Text = txtNo.Text & "."
 
            txtStatus.Text = .Fields("ng_name").Value
            txt00.Text = .Fields("C00").Value
            txt01.Text = .Fields("C01").Value
            txt02.Text = .Fields("C02").Value
            txt03.Text = .Fields("C03").Value
            txt04.Text = .Fields("C04").Value
            txt05.Text = .Fields("C05").Value
            txt06.Text = .Fields("C06").Value
            txt07.Text = .Fields("C07").Value
            txtShift3.Text = .Fields("Shift_3").Value

            txt08.Text = .Fields("C08").Value
            txt09.Text = .Fields("C09").Value
            txt10.Text = .Fields("C10").Value
            txt11.Text = .Fields("C11").Value
            txt12.Text = .Fields("C12").Value
            txt13.Text = .Fields("C13").Value
            txt14.Text = .Fields("C14").Value
            txt15.Text = .Fields("C15").Value
            txtShift1.Text = .Fields("Shift_1").Value

            txt16.Text = .Fields("C16").Value
            txt17.Text = .Fields("C17").Value
            txt18.Text = .Fields("C18").Value
            txt19.Text = .Fields("C19").Value
            txt20.Text = .Fields("C20").Value
            txt21.Text = .Fields("C21").Value
            txt22.Text = .Fields("C22").Value
            txt23.Text = .Fields("C23").Value
            txtShift2.Text = .Fields("Shift_2").Value
        End If
    End With

End Sub


