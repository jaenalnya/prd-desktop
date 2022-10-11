VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptSkillGeneral 
   Caption         =   "Skill Matrik General"
   ClientHeight    =   14025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   28680
   Icon            =   "RptSkillGeneral.dsx":0000
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   50588
   _ExtentY        =   24739
   SectionData     =   "RptSkillGeneral.dsx":617A
End
Attribute VB_Name = "RptSkillGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim qSQL As String

Private Sub ActiveReport_DataInitialize()
On Error Resume Next
qSQL = "select a.*,b.nik,b.nama_karyawan,b.departement,b.positions,"
qSQL = qSQL & " a.sys_plant_id , c.prod_skill_general_id, c.prod_skill_general_list_id, c.Description, skill"
qSQL = qSQL & " from prod_skill_generals a"
qSQL = qSQL & " INNER JOIN (select emp.id, emp.nik,emp.name as nama_karyawan,dep.name as departement,"
        qSQL = qSQL & " pos.name as positions from hrd_employees emp"
        qSQL = qSQL & " inner join sys_departments dep on emp.sys_department_id = dep.id"
        qSQL = qSQL & " inner join hrd_positions pos on emp.hrd_position_id = pos.id) b on a.hrd_employee_id = b.id"
qSQL = qSQL & " INNER JOIN (select skl_itm.id,skl_itm.prod_skill_general_id, skl_itm.prod_skill_general_list_id,"
        qSQL = qSQL & " skl_lst.description,skl_itm.skill from prod_skill_general_items skl_itm"
        qSQL = qSQL & " inner join prod_skill_general_lists skl_lst on skl_itm.prod_skill_general_list_id = skl_lst.id) c"
        qSQL = qSQL & " on a.id = c.prod_skill_general_id"
qSQL = qSQL & " where b.nik = '" & Mid(ACTIVE_USER.USERNAME, 2, 7) & "' AND a.status_doc = 'active'"

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
    .txtPosisi.DataField = "positions"
    .txtDept.DataField = "departement"
    .txtperiode.DataField = ""
    .txtNextEvaluasi.DataField = ""
    .txtSkill.DataField = "skill"
    .txtDeskription.DataField = "Description"

        
End With
End Sub

Private Sub Detail_Format()
On Error Resume Next
    With DTRpt.Recordset
        If Not .EOF Then
            txtNo.Text = Val(txtNo.Text) + 1
            txtNo.Text = txtNo.Text & "."
            
            If RS_PRINT!skill >= 0 And RS_PRINT!skill <= 9 Then
                Img_Skill.Picture = LoadPicture(App.Path & "\image\point_0.jpg")
    
    
            ElseIf RS_PRINT!skill >= 10 And RS_PRINT!skill <= 35 Then
                Img_Skill.Picture = LoadPicture(App.Path & "\image\point_1.jpg")
            
        
            ElseIf RS_PRINT!skill >= 36 And RS_PRINT!skill <= 60 Then
                Img_Skill.Picture = LoadPicture(App.Path & "\image\point_2.jpg")
        
            ElseIf RS_PRINT!skill >= 61 And RS_PRINT!skill <= 85 Then
                Img_Skill.Picture = LoadPicture(App.Path & "\image\point_3.jpg")
                
            ElseIf RS_PRINT!skill >= 86 And RS_PRINT!skill <= 100 Then
                Img_Skill.Picture = LoadPicture(App.Path & "\image\point_4.jpg")
            End If
            
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


