VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptHarian 
   Caption         =   "Laporan Harian"
   ClientHeight    =   9690
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   20370
   Icon            =   "rptHarian.dsx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   35930
   _ExtentY        =   17092
   SectionData     =   "rptHarian.dsx":617A
End
Attribute VB_Name = "rptHarian"
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
        .txtDataGroup.DataField = "DataGroup"
        .txtTanggal.DataField = "Tanggal"
        .txtDTanggal.DataField = "Tanggal"
        .txtJam.DataField = "Jam"
        .txtNamaBarang.DataField = "NamaBarang"
        .txtBerat.DataField = "Berat"
        .txtTotalBerat.DataField = "Berat"

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






