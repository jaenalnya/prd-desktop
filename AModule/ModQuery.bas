Attribute VB_Name = "ModQuery"
Option Explicit
Public p_plant_mark                  As String
Public p_sys_plant                   As String
Public p_sys_plant_id                As String
Public p_prod_machine_id             As String
Public p_machine_no                  As String
Public p_machine_name                As String
Public p_tonnage                     As String
Public p_mkt_customer_id             As String
Public p_customer_name               As String
Public p_eng_product_1               As String
Public p_prod_name_1                 As String
Public p_int_part_1                  As String
Public p_cycle_time_1                As String
Public p_cavity_1                    As String
Public p_weight_gr_1                 As String
Public p_weight_runner_gr_1          As String
Public p_status_prod_1               As Boolean

Public p_eng_product_2               As String
Public p_prod_name_2                 As String
Public p_int_part_2                  As String
Public p_status_prod_2               As Boolean

Public p_eng_product_3               As String
Public p_prod_name_3                 As String
Public p_int_part_3                  As String
Public p_status_prod_3               As Boolean

Public p_eng_product_4               As String
Public p_prod_name_4                 As String
Public p_int_part_4                  As String
Public p_status_prod_4               As Boolean

Public p_machine_status              As String
Public NoMesin                       As String
Public Sub LoadProduct()
On Error Resume Next

    Dim Rs As New ADODB.Recordset
    Rs.CursorLocation = adUseClient

    Dim sSQL        As String

    sSQL = "SELECT * FROM sip_production.view_prod_running_products WHERE plant_mark = '" & p_plant_mark & "' and machine_no = '" & NoMesin & "'"

    If Rs.State = adStateOpen Then Rs.Close
    Rs.Open sSQL, CN, adOpenDynamic, adLockPessimistic
    
    If Rs.RecordCount > 0 Then
        With Rs
            'p_plant_mark = .Fields("plant_mark")
            p_prod_machine_id = .Fields("prod_machine_id")
            p_sys_plant_id = .Fields("sys_plant_id")
            p_machine_no = .Fields("machine_no")
            p_machine_name = .Fields("machine_name")
            p_tonnage = .Fields("tonnage")
            p_mkt_customer_id = .Fields("mkt_customer_id")
            p_customer_name = .Fields("customer_name")
            p_eng_product_1 = .Fields("eng_product_1")
            p_prod_name_1 = .Fields("prod_name_1")
            p_int_part_1 = .Fields("int_part_1")
            p_cycle_time_1 = IIf(IsNull(.Fields("cycle_time_ia_1")), "0", .Fields("cycle_time_ia_1"))
            p_cavity_1 = IIf(IsNull(.Fields("cavity_1")), "0", .Fields("cavity_1"))
            p_weight_gr_1 = IIf(IsNull(.Fields("weight_gr_1")), "0", .Fields("weight_gr_1"))
            p_weight_runner_gr_1 = IIf(IsNull(.Fields("weight_runner_gr_1")), "0", .Fields("weight_runner_gr_1"))
            
            p_eng_product_2 = IIf(IsNull(.Fields("eng_product_2")), "", .Fields("eng_product_2"))
            p_prod_name_2 = IIf(IsNull(.Fields("prod_name_2")), "", .Fields("prod_name_2"))
            p_int_part_2 = IIf(IsNull(.Fields("int_part_2")), "", .Fields("int_part_2"))

            p_eng_product_3 = IIf(IsNull(.Fields("eng_product_3")), "", .Fields("eng_product_3"))
            p_prod_name_3 = IIf(IsNull(.Fields("prod_name_3")), "", .Fields("prod_name_3"))
            p_int_part_3 = IIf(IsNull(.Fields("int_part_3")), "", .Fields("int_part_3"))
            
            p_eng_product_4 = IIf(IsNull(.Fields("eng_product_4")), "", .Fields("eng_product_4"))
            p_prod_name_4 = IIf(IsNull(.Fields("prod_name_4")), "", .Fields("prod_name_4"))
            p_int_part_4 = IIf(IsNull(.Fields("int_part_4")), "", .Fields("int_part_4"))
            
            
            p_machine_status = .Fields("machine_status")
        End With

     End If
     
        If p_eng_product_1 = "" Then
            p_status_prod_1 = False
        Else
            p_status_prod_1 = True
        End If
        
        If p_eng_product_2 = "" Then
            p_status_prod_2 = False
        Else
            p_status_prod_2 = True
        End If

        If p_eng_product_3 = "" Then
            p_status_prod_3 = False
        Else
            p_status_prod_3 = True
        End If

        If p_eng_product_4 = "" Then
            p_status_prod_4 = False
        Else
            p_status_prod_4 = True
        End If
        
        
    
    Set Rs = Nothing
   
End Sub

