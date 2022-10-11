VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptPerBarang 
   Caption         =   "LAPORAN PER BARANG"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20250
   Icon            =   "RptPerBarang.dsx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   _ExtentX        =   35719
   _ExtentY        =   18203
   SectionData     =   "RptPerBarang.dsx":169B2
End
Attribute VB_Name = "RptPerBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_DataInitialize()
With Me
    .lblDate.Caption = Now
    .lblCompany.Caption = ACTIVE_COMPANY.Perusahaan
    .lblAddress.Caption = ACTIVE_COMPANY.Alamat
    .lblContact.Caption = ACTIVE_COMPANY.NoTelepon

    .txtIDTransaksi.DataField = "IDTransaksi"
    .txtShift.DataField = "Shift"
    .txtTanggal.DataField = "Tanggal"
    .txtJam.DataField = "Jam"
    .txtNamabarang.DataField = "NamaBarang"
    .txtBerat.DataField = "Berat"
    .txtTotalBerat.DataField = "Berat"
    .txtGTotalBerat.DataField = "Berat"
    .txtDataGroup.DataField = "DataGroup"
End With
End Sub

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

