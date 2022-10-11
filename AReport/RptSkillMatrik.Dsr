VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptSkillMatrik 
   Caption         =   "Skill Matrik"
   ClientHeight    =   9600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13020
   Icon            =   "RptSkillMatrik.dsx":0000
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   22966
   _ExtentY        =   16933
   SectionData     =   "RptSkillMatrik.dsx":617A
End
Attribute VB_Name = "RptSkillMatrik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim qSQL As String

Private Sub ActiveReport_DataInitialize()
On Error Resume Next
qSQL = "select a.*,b.nik,b.nama_karyawan,b.departement,b.positions,"
qSQL = qSQL & " c.prod_skill_product_id,c.eng_product_id,c.internal_part_id,c.product_name,"
qSQL = qSQL & " c.pg,c.cv,c.rw,c.pl,c.ng,c.result,c.`status`,"
qSQL = qSQL & " c.created_at , c.created_by, c.created_name, c.approve_1_at, c.approve_1_by, c.app1_name"
qSQL = qSQL & " from prod_skill_products a"
qSQL = qSQL & " INNER JOIN (select emp.id, emp.nik,emp.name as nama_karyawan,dep.name as departement,"
        qSQL = qSQL & " pos.name as positions from sip_production.hrd_employees emp"
        qSQL = qSQL & " inner join sip_production.sys_departments dep on emp.sys_department_id = dep.id"
        qSQL = qSQL & " inner join sip_production.hrd_positions pos on emp.hrd_position_id = pos.id) b on a.hrd_employee_id = b.id"
qSQL = qSQL & " INNER JOIN (SELECT skill.id,skill.prod_skill_product_id,skill.eng_product_id,prd.internal_part_id,"
        qSQL = qSQL & " prd.name as product_name,skill.pg,skill.cv,skill.rw,skill.pl,skill.ng,skill.result,skill.`status`,"
        qSQL = qSQL & " skill.created_at,skill.created_by,crt_sys.name as created_name,"
        qSQL = qSQL & " skill.approve_1_at,skill.approve_1_by,app1_sys.name as app1_name"
        qSQL = qSQL & " FROM sip_production.prod_skill_product_items skill"
        qSQL = qSQL & " LEFT JOIN sip_production.eng_products prd on skill.eng_product_id = prd.id"
        qSQL = qSQL & " LEFT JOIN sip_production.sys_accounts crt_sys on skill.created_by = crt_sys.id"
        qSQL = qSQL & " LEFT JOIN sip_production.sys_accounts app1_sys on skill.approve_1_by = app1_sys.id) c"
qSQL = qSQL & " on a.id = c.prod_skill_product_id where b.nik = '" & Mid(ACTIVE_USER.USERNAME, 2, 7) & "' and c.status = 'active'"

Set RS_PRINT = New ADODB.Recordset
If RS_PRINT.State = adStateOpen Then RS_PRINT.Close
RS_PRINT.Open qSQL, CN, adOpenDynamic, adLockPessimistic
With Me
    .DTRpt.Recordset = RS_PRINT
    .lblDate.Caption = Now
    .lblCompany.Caption = ACTIVE_COMPANY.Perusahaan
    .lblAlamat.Caption = ACTIVE_COMPANY.Alamat
    .txtNik.DataField = "Nik"
    .txtNama.DataField = "nama_karyawan"
    .txtDept.DataField = "departement"
    .txtPosisi.DataField = "positions"
    .txtProduk.DataField = "product_name"
    .txtPg.DataField = "pg"
    .txtCv.DataField = "cv"
    .txtRw.DataField = "rw"
    .txtPl.DataField = "pl"
    .txtng.DataField = "ng"
    .txtHasil.DataField = "result"
    .txtTrainner.DataField = "created_name"
    .txtLeader.DataField = "app1_name"
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


