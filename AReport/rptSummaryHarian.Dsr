VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptSummaryHarian 
   Caption         =   "Laporan Total Harian"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13845
   Icon            =   "rptSummaryHarian.dsx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   24421
   _ExtentY        =   11483
   SectionData     =   "rptSummaryHarian.dsx":169B2
End
Attribute VB_Name = "rptSummaryHarian"
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

    .txtHTanggal.DataField = "Tanggal"
    .txtDTanggal.DataField = "Tanggal"
    .txtIndikator.DataField = "Indikator"
    .txtNamaBarang.DataField = "NamaBarang"
    .txtBerat.DataField = "TotalBerat"
    .txtTotalBerat.DataField = "TotalBerat"
    .txtGTotal.DataField = "TotalBerat"
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



