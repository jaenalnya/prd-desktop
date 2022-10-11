VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptProduksi 
   Caption         =   "PRD - RptProduksi (ActiveReport)"
   ClientHeight    =   12375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   Icon            =   "RptProduksi.dsx":0000
   WindowState     =   2  'Maximized
   _ExtentX        =   40217
   _ExtentY        =   21828
   SectionData     =   "RptProduksi.dsx":617A
End
Attribute VB_Name = "RptProduksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_DataInitialize()
With Me

    .lblDate.Caption = Now
    .lblCompany.Caption = ACTIVE_COMPANY.Perusahaan
    .lblAlamat.Caption = ACTIVE_COMPANY.Alamat
    '.lblContact.Caption = ACTIVE_COMPANY.NoTelepon
    .txtMachineID.DataField = "prod_machine_id"
    .txtNamaMesin.DataField = "machine_name"
    .txtNoMesin.DataField = "number"
    .txtTonage.DataField = "tonnage"
    .txtProduct.DataField = "product_name"
    .txtProdID.DataField = "eng_product_id"
    .txtInternalID.DataField = "internal_part_id"
    .txtCustomerID.DataField = "customer_part_number"
    .txtCapity.DataField = "cavity"
    .txtCycleTime.DataField = "cycle_time_ia"
    .txtShot.DataField = "target_shot"
    .txtWeight.DataField = "weight_gr"
    .txtRunner.DataField = "weight_runner_gr"
    .txtdate.DataField = "period_shift"
    .txtMaterial.DataField = "material_name"
    
End With

End Sub


Private Sub QueryShot()

    Dim strSQL  As String
    Dim RS_SHOT As New ADODB.Recordset

    strSQL = "select plant_mark,prod_machine_id,mkt_customer_id,eng_product_id,period_shift, status, "
        strSQL = strSQL & " C08,C09,C10,C11,C12,C13,C14,C15,(C08+C09+C10+C11+C12+C13+C14+C15) as Shift_1,"
        strSQL = strSQL & " C16,C17,C18,C19,C20,C21,C22,C23,(C16+C17+C18+C19+C20+C21+C22+C23) as Shift_2,"
        strSQL = strSQL & " C00,C01,C02,C03,C04,C05,C06,C07,(C00+C01+C02+C03+C04+C05+C06+C07) as Shift_3"
        strSQL = strSQL & " From (select plant_mark,prod_machine_id,mkt_customer_id,eng_product_id,period_shift, status,"
        strSQL = strSQL & " SUM(IF(period_hour = '08', jumlah, 0)) AS 'C08', SUM(IF(period_hour = '09', jumlah, 0)) AS 'C09',"
        strSQL = strSQL & " SUM(IF(period_hour = '10', jumlah, 0)) AS 'C10', SUM(IF(period_hour = '11', jumlah, 0)) AS 'C11',"
        strSQL = strSQL & " SUM(IF(period_hour = '12', jumlah, 0)) AS 'C12', SUM(IF(period_hour = '13', jumlah, 0)) AS 'C13',"
        strSQL = strSQL & " SUM(IF(period_hour = '14', jumlah, 0)) AS 'C14', SUM(IF(period_hour = '15', jumlah, 0)) AS 'C15',"
        strSQL = strSQL & " SUM(IF(period_hour = '16', jumlah, 0)) AS 'C16', SUM(IF(period_hour = '17', jumlah, 0)) AS 'C17',"
        strSQL = strSQL & " SUM(IF(period_hour = '18', jumlah, 0)) AS 'C18', SUM(IF(period_hour = '19', jumlah, 0)) AS 'C19',"
        strSQL = strSQL & " SUM(IF(period_hour = '20', jumlah, 0)) AS 'C20', SUM(IF(period_hour = '21', jumlah, 0)) AS 'C21',"
        strSQL = strSQL & " SUM(IF(period_hour = '22', jumlah, 0)) AS 'C22', SUM(IF(period_hour = '23', jumlah, 0)) AS 'C23',"
        strSQL = strSQL & " SUM(IF(period_hour = '00', jumlah, 0)) AS 'C00', SUM(IF(period_hour = '01', jumlah, 0)) AS 'C01',"
        strSQL = strSQL & " SUM(IF(period_hour = '02', jumlah, 0)) AS 'C02', SUM(IF(period_hour = '03', jumlah, 0)) AS 'C03',"
        strSQL = strSQL & " SUM(IF(period_hour = '04', jumlah, 0)) AS 'C04', SUM(IF(period_hour = '05', jumlah, 0)) AS 'C05',"
        strSQL = strSQL & " SUM(IF(period_hour = '06', jumlah, 0)) AS 'C06', SUM(IF(period_hour = '07', jumlah, 0)) AS 'C07'"
        strSQL = strSQL & " From (select a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift,a.period_hour,"
                strSQL = strSQL & " sum(a.counter_ok) as jumlah,'1. SHOT' as status from sip_production.prod_runnings a"
                strSQL = strSQL & " group by a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift,a.period_hour"
        strSQL = strSQL & " Union select a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift,a.period_hour,"
                strSQL = strSQL & " sum(a.counter_ng) as jumlah,'3. NG PROD' as status  from sip_production.prod_data_ngs a where status = 'active' "
                strSQL = strSQL & " group by a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift,a.period_hour"
        strSQL = strSQL & " Union select a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift,a.period_hour,"
                strSQL = strSQL & " sum(a.counter_ok * b.cavity) as jumlah,'2. GROSS' as status from sip_production.prod_runnings a"
                strSQL = strSQL & " inner join sip_production.eng_products b on a.eng_product_id = b.id group by a.plant_mark,a.prod_machine_id,a.mkt_customer_id,"
                strSQL = strSQL & " a.eng_product_id , a.period_shift, a.period_hour"
        strSQL = strSQL & " Union select xx.plant_mark,xx.prod_machine_id,xx.mkt_customer_id,xx.eng_product_id,xx.period_shift,xx.period_hour,"
                strSQL = strSQL & " (xx.gross_produksi-ifnull(yy.ng,0)) as net_produksi,'4. NET_PROD' as status from (select a.plant_mark,a.prod_machine_id,"
                strSQL = strSQL & " a.mkt_customer_id,a.eng_product_id,a.period_shift,a.period_hour, a.counter_ok as shot, sum(a.counter_ok * b.cavity) as gross_produksi"
                strSQL = strSQL & " from sip_production.prod_runnings a inner join sip_production.eng_products b on a.eng_product_id = b.id"
                strSQL = strSQL & " group by a.plant_mark,a.prod_machine_id,a.mkt_customer_id, a.eng_product_id,a.period_shift,a.period_hour) xx"
        strSQL = strSQL & " left join (select a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift,a.period_hour,"
                strSQL = strSQL & " sum(a.counter_ng) as ng from sip_production.prod_data_ngs a where status = 'active' group by a.plant_mark,a.prod_machine_id,a.mkt_customer_id,"
                strSQL = strSQL & " a.eng_product_id,a.period_shift,a.period_hour) yy on xx.plant_mark = yy.plant_mark"
                strSQL = strSQL & " and xx.prod_machine_id = yy.prod_machine_id and xx.mkt_customer_id = yy.mkt_customer_id"
                strSQL = strSQL & " and xx.eng_product_id = yy.eng_product_id and xx.period_shift = yy.period_shift and xx.period_hour = yy.period_hour ) as x"
                strSQL = strSQL & " group by plant_mark,prod_machine_id,mkt_customer_id,eng_product_id,period_shift, status) data_produksi"
        strSQL = strSQL & " where plant_mark = '" & p_plant_mark & "'"
        strSQL = strSQL & " and prod_machine_id = '" & txtMachineID.Text & "'"
        strSQL = strSQL & " and eng_product_id = '" & txtProdID.Text & "'"
        strSQL = strSQL & " and period_shift = '" & Format(txtdate.Text, "yyyy-mm-dd") & "'"

    Set RS_SHOT = New ADODB.Recordset
    If RS_SHOT.State = adStateOpen Then RS_SHOT.Close
    RS_SHOT.Open strSQL, CN, adOpenDynamic, adLockPessimistic
    ' memanggil sub report
    Set SubReport3.object = New RptSubShot
    With SubReport3.object.DTRpt
        .Recordset = RS_SHOT
        .Source = strSQL
    End With
    
End Sub

Private Sub QueryNG()
    Dim strSQL  As String
    Dim RS_SUB As New ADODB.Recordset

    strSQL = "select plant_mark,prod_machine_id,mkt_customer_id,eng_product_id,period_shift,prod_ng_id, ng_name,  "
    strSQL = strSQL & " C08,C09,C10,C11,C12,C13,C14,C15,(C08+C09+C10+C11+C12+C13+C14+C15) as Shift_1,"
    strSQL = strSQL & " C16,C17,C18,C19,C20,C21,C22,C23,(C16+C17+C18+C19+C20+C21+C22+C23) as Shift_2,"
    strSQL = strSQL & " C00,C01,C02,C03,C04,C05,C06,C07,(C00+C01+C02+C03+C04+C05+C06+C07) as Shift_3"
    strSQL = strSQL & " From ( select plant_mark,prod_machine_id,mkt_customer_id,eng_product_id,period_shift,prod_ng_id, ng_name,"
    strSQL = strSQL & " SUM(IF(period_hour = '08', jumlah, 0)) AS 'C08', SUM(IF(period_hour = '09', jumlah, 0)) AS 'C09',"
    strSQL = strSQL & " SUM(IF(period_hour = '10', jumlah, 0)) AS 'C10', SUM(IF(period_hour = '11', jumlah, 0)) AS 'C11',"
    strSQL = strSQL & " SUM(IF(period_hour = '12', jumlah, 0)) AS 'C12', SUM(IF(period_hour = '13', jumlah, 0)) AS 'C13',"
    strSQL = strSQL & " SUM(IF(period_hour = '14', jumlah, 0)) AS 'C14', SUM(IF(period_hour = '15', jumlah, 0)) AS 'C15',"
    strSQL = strSQL & " SUM(IF(period_hour = '16', jumlah, 0)) AS 'C16', SUM(IF(period_hour = '17', jumlah, 0)) AS 'C17',"
    strSQL = strSQL & " SUM(IF(period_hour = '18', jumlah, 0)) AS 'C18', SUM(IF(period_hour = '19', jumlah, 0)) AS 'C19',"
    strSQL = strSQL & " SUM(IF(period_hour = '20', jumlah, 0)) AS 'C20', SUM(IF(period_hour = '21', jumlah, 0)) AS 'C21',"
    strSQL = strSQL & " SUM(IF(period_hour = '22', jumlah, 0)) AS 'C22', SUM(IF(period_hour = '23', jumlah, 0)) AS 'C23',"
    strSQL = strSQL & " SUM(IF(period_hour = '00', jumlah, 0)) AS 'C00', SUM(IF(period_hour = '01', jumlah, 0)) AS 'C01',"
    strSQL = strSQL & " SUM(IF(period_hour = '02', jumlah, 0)) AS 'C02', SUM(IF(period_hour = '03', jumlah, 0)) AS 'C03',"
    strSQL = strSQL & " SUM(IF(period_hour = '04', jumlah, 0)) AS 'C04', SUM(IF(period_hour = '05', jumlah, 0)) AS 'C05',"
    strSQL = strSQL & " SUM(IF(period_hour = '06', jumlah, 0)) AS 'C06', SUM(IF(period_hour = '07', jumlah, 0)) AS 'C07'"
    strSQL = strSQL & " From (select a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift,"
    strSQL = strSQL & " period_hour, a.prod_ng_id,b.name as ng_name, sum(a.counter_ng) as jumlah"
    strSQL = strSQL & " from sip_production.prod_data_ngs a inner join sip_production.prod_ngs b on a.prod_ng_id = b.id"
    strSQL = strSQL & " where a.status = 'active'"
    strSQL = strSQL & " group by a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift, a.period_hour,a.prod_ng_id,b.name ) as X"
    strSQL = strSQL & " group by plant_mark,prod_machine_id,mkt_customer_id,eng_product_id,period_shift,prod_ng_id, ng_name ) xx "
    strSQL = strSQL & " where plant_mark = '" & p_plant_mark & "'"
    strSQL = strSQL & " and prod_machine_id = '" & txtMachineID.Text & "'"
    strSQL = strSQL & " and eng_product_id = '" & txtProdID.Text & "'"
    strSQL = strSQL & " and period_shift = '" & Format(txtdate.Text, "yyyy-mm-dd") & "'"

    Set RS_SUB = New ADODB.Recordset
    If RS_SUB.State = adStateOpen Then RS_SUB.Close
    RS_SUB.Open strSQL, CN, adOpenDynamic, adLockPessimistic
    ' memanggil sub report
    Set SubReport1.object = New RptSubProduksi
    With SubReport1.object.DTRpt
        .Recordset = RS_SUB
        .Source = strSQL
    End With
End Sub
Private Sub QuerySubIdle()
    Dim ssSql  As String
    Dim RS_SUBidle As New ADODB.Recordset

    ssSql = "select a.hrd_employee_id,d.name as karyawan,a.plant_mark,a.prod_machine_id,b.number,b.name as machine_name,"
    ssSql = ssSql & " a.mkt_customer_id,a.eng_product_1,a.eng_product_2,a.prod_idletime_id,a.period_shift,a.start_idle,a.end_idle,a.idle_time,"
    ssSql = ssSql & " a.description,c.name as idle_name,c.description as kode"
    ssSql = ssSql & " from sip_production.prod_machine_idles a"
    ssSql = ssSql & " INNER JOIN sip_production.prod_machines b ON a.prod_machine_id = b.id"
    ssSql = ssSql & " INNER JOIN sip_production.prod_idletimes c ON a.prod_idletime_id = c.id"
    ssSql = ssSql & " LEFT JOIN sip_production.hrd_employees d ON a.hrd_employee_id = d.id"
    ssSql = ssSql & " where a.plant_mark = '" & p_plant_mark & "'"
    ssSql = ssSql & " and a.prod_machine_id = '" & txtMachineID.Text & "'"
    ssSql = ssSql & " and a.period_shift = '" & Format(txtdate.Text, "yyyy-mm-dd") & "'"

    Set RS_SUBidle = New ADODB.Recordset
    If RS_SUBidle.State = adStateOpen Then RS_SUBidle.Close
    RS_SUBidle.Open ssSql, CN, adOpenDynamic, adLockPessimistic
    ' memanggil sub report
    Set SubReport2.object = New RptSubIdle
    With SubReport2.object.DTRpt
        .Recordset = RS_SUBidle
        .Source = ssSql
    End With
    
End Sub

Private Sub QueryOpt()
    Dim ssSql  As String
    Dim RS_OPT As New ADODB.Recordset

    ssSql = "SELECT A.plant_mark,A.prod_machine_id,D.number AS machine_no, A.operator_1,B.nik AS nik_operator_1, B.name AS name_operator_1,"
    ssSql = ssSql & " A.operator_2,C.nik AS nik_operator_2, C.name AS name_operator_2"
    ssSql = ssSql & " FROM prod_runnings A"
    ssSql = ssSql & " LEFT JOIN hrd_employees B ON A.operator_1 = B.id"
    ssSql = ssSql & " LEFT JOIN hrd_employees C ON A.operator_2 = C.id"
    ssSql = ssSql & " LEFT JOIN prod_machines D ON A.prod_machine_id = D.id"
    ssSql = ssSql & " where A.plant_mark = '" & p_plant_mark & "'"
    ssSql = ssSql & " and A.prod_machine_id = '" & txtMachineID.Text & "'"
    ssSql = ssSql & " and A.period_shift = '" & Format(txtdate.Text, "yyyy-mm-dd") & "'"
    ssSql = ssSql & " GROUP BY A.plant_mark,A.prod_machine_id, A.operator_1,A.operator_2"

    Set RS_OPT = New ADODB.Recordset
    If RS_OPT.State = adStateOpen Then RS_OPT.Close
    RS_OPT.Open ssSql, CN, adOpenDynamic, adLockPessimistic
    ' memanggil sub report
    Set SubReport4.object = New RptOperator
    With SubReport4.object.DTRpt
        .Recordset = RS_OPT
        .Source = ssSql
    End With
    
End Sub

Private Sub QueryOK()
    Dim ssSql  As String
    Dim RS_OK As New ADODB.Recordset

    ssSql = "select a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift,a.shift,"
    ssSql = ssSql & " SUM(IF(product_status = 'ok', qty, 0)) AS  'ok',"
    ssSql = ssSql & " SUM(IF(product_status = 'sisa', qty, 0)) AS  'sisa',"
    ssSql = ssSql & " SUM(IF(product_status = 'hold', qty, 0)) AS  'hold',"
    ssSql = ssSql & " (SUM(IF(product_status = 'ok', qty, 0))+SUM(IF(product_status = 'sisa', qty, 0))+SUM(IF(product_status = 'hold', qty, 0))) as Sub_total"
    ssSql = ssSql & " from prod_result_logs a"
    ssSql = ssSql & " where a.`status` = 'active'"
    ssSql = ssSql & " and a.plant_mark = '" & p_plant_mark & "'"
    ssSql = ssSql & " and a.prod_machine_id = '" & txtMachineID.Text & "'"
    ssSql = ssSql & " and a.eng_product_id = '" & txtProdID.Text & "'"
    ssSql = ssSql & " and a.period_shift = '" & Format(txtdate.Text, "yyyy-mm-dd") & "'"
    ssSql = ssSql & " group by a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift,a.shift"
    ssSql = ssSql & " order by a.period_shift desc, a.shift Asc"
    
    Set RS_OK = New ADODB.Recordset
    If RS_OK.State = adStateOpen Then RS_OK.Close
    RS_OK.Open ssSql, CN, adOpenDynamic, adLockPessimistic
    ' memanggil sub report
    Set SubReport5.object = New RptSubSPB2
    With SubReport5.object.DTRpt
        .Recordset = RS_OK
        .Source = ssSql
    End With
    
End Sub

Private Sub QuerySummary()
    Dim ssSql  As String
    Dim RS_OK As New ADODB.Recordset

    ssSql = "select X.plant_mark,X.prod_machine_id,X.mkt_customer_id,X.eng_product_id, X.period_shift,X.shift,X.total,Y.cavity,(X.total * Y.cavity) as gross,Z.total_ng,"
    ssSql = ssSql & " ((X.total * Y.cavity) - Z.total_ng) as net_produksi,round(((X.total * Y.cavity) - Z.total_ng)/(X.total * Y.cavity) * 100) as prod_yield"
    ssSql = ssSql & " From"
            ssSql = ssSql & " (select aa.plant_mark,aa.prod_machine_id,aa.mkt_customer_id,aa.eng_product_id, aa.period_shift,sum(aa.counter_ok) as total,'1' as shift from prod_runnings aa"
            ssSql = ssSql & " where aa.period_hour between '08' and '15'"
            ssSql = ssSql & " group by aa.plant_mark,aa.prod_machine_id,aa.mkt_customer_id,aa.eng_product_id, aa.period_shift"
            ssSql = ssSql & " Union"
            ssSql = ssSql & " select aa.plant_mark,aa.prod_machine_id,aa.mkt_customer_id,aa.eng_product_id, aa.period_shift,sum(aa.counter_ok) as total,'2' as shift from prod_runnings aa"
            ssSql = ssSql & " where aa.period_hour between '16' and '23'"
            ssSql = ssSql & " group by aa.plant_mark,aa.prod_machine_id,aa.mkt_customer_id,aa.eng_product_id, aa.period_shift"
            ssSql = ssSql & " Union"
            ssSql = ssSql & " select aa.plant_mark,aa.prod_machine_id,aa.mkt_customer_id,aa.eng_product_id, aa.period_shift,sum(aa.counter_ok) as total,'3' as shift from prod_runnings aa"
            ssSql = ssSql & " where aa.period_hour between '00' and '07'"
            ssSql = ssSql & " group by aa.plant_mark,aa.prod_machine_id,aa.mkt_customer_id,aa.eng_product_id, aa.period_shift) as X"
    ssSql = ssSql & " Left Join"
            ssSql = ssSql & " eng_products Y on X.eng_product_id = Y.id"
    ssSql = ssSql & " Left Join"
            ssSql = ssSql & " (select a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift,'1' as shift,sum(a.counter_ng) as total_ng from prod_data_ngs a"
            ssSql = ssSql & " WHERE a.period_hour between '08' and '15'"
            ssSql = ssSql & " group by a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift"
            ssSql = ssSql & " Union"
            ssSql = ssSql & " select a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift,'2' as shift,sum(a.counter_ng) as total_ng from prod_data_ngs a"
            ssSql = ssSql & " WHERE a.period_hour between '16' and '23'"
            ssSql = ssSql & " group by a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift"
            ssSql = ssSql & " Union"
            ssSql = ssSql & " select a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift,'3' as shift,sum(a.counter_ng) as total_ng from prod_data_ngs a"
            ssSql = ssSql & " WHERE a.period_hour between '00' and '07'"
            ssSql = ssSql & " group by a.plant_mark,a.prod_machine_id,a.mkt_customer_id,a.eng_product_id,a.period_shift) as Z"
ssSql = ssSql & " on X.plant_mark = Z.plant_mark and X.prod_machine_id = Z.prod_machine_id and X.mkt_customer_id = Z.mkt_customer_id and"
ssSql = ssSql & " X.eng_product_id = Z.eng_product_id and X.period_shift = Z.period_shift and X.shift = Z.shift"
            ssSql = ssSql & " where X.plant_mark = '" & p_plant_mark & "'"
            ssSql = ssSql & " and X.prod_machine_id = '" & txtMachineID.Text & "'"
            ssSql = ssSql & " and X.eng_product_id = '" & txtProdID.Text & "'"
            ssSql = ssSql & " and X.period_shift = '" & Format(txtdate.Text, "yyyy-mm-dd") & "'"
ssSql = ssSql & " order by X.period_shift desc,X.prod_machine_id asc,X.shift ASC"
    
    Set RS_OK = New ADODB.Recordset
    If RS_OK.State = adStateOpen Then RS_OK.Close
    RS_OK.Open ssSql, CN, adOpenDynamic, adLockPessimistic
    ' memanggil sub report
    Set SubReport6.object = New RptSubSummary
    With SubReport6.object.DTRpt
        .Recordset = RS_OK
        .Source = ssSql
    End With
    
End Sub


Private Sub GroupHeader1_Format()
    Call QueryShot
End Sub

Private Sub GroupHeader2_Format()
    Call QueryOK
    Call QuerySummary
    
End Sub

Private Sub GroupHeader3_Format()
    Call QueryNG
End Sub

Private Sub GroupHeader4_Format()
    Call QuerySubIdle
    Call QueryOpt
End Sub
