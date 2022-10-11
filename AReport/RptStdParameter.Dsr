VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} RptStdParameter 
   Caption         =   "Standard Parameter"
   ClientHeight    =   9870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17640
   Icon            =   "RptStdParameter.dsx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   31115
   _ExtentY        =   17410
   SectionData     =   "RptStdParameter.dsx":617A
End
Attribute VB_Name = "RptStdParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim qSQL As String

Private Sub ActiveReport_DataInitialize()
With Me
    .DTRpt.Recordset = RS_PRINT
    .lblDate.Caption = Now
    .lblCompany.Caption = ACTIVE_COMPANY.Perusahaan
    .lblAlamat.Caption = ACTIVE_COMPANY.Alamat
    .txtnomc.DataField = "machine_number"
    .txtmachine_name.DataField = "machine_name"
    .txtnumber.DataField = "number"
    .txtdate.DataField = "date"
    .txtrev.DataField = "rev"
    .txtrev_date.DataField = "rev_date"
    
    .txtcustomer_name.DataField = "customer_name"
    .txtcustomer_part_name.DataField = "customer_part_name"
    .txtcustomer_part_number.DataField = "customer_part_number"
    
    .txtcolor_name.DataField = "color_name"
    .txtmaterial_name.DataField = "material_name"
    
    
    
    .txtinjectrol_ph_4.DataField = "injectrol_ph_4"
    .txtinjectrol_ph_3.DataField = "injectrol_ph_3"
    .txtinjectrol_ph_2.DataField = "injectrol_ph_2"
    .txtinjectrol_ph_1.DataField = "injectrol_ph_1"
    .txtinjectrol_vh_1.DataField = "injectrol_vh_1"
    .txtinjectrol_vi_5.DataField = "injectrol_vi_5"
    .txtinjectrol_vi_4.DataField = "injectrol_vi_4"
    .txtinjectrol_vi_3.DataField = "injectrol_vi_3"
    .txtinjectrol_vi_2.DataField = "injectrol_vi_2"
    .txtinjectrol_vi_1.DataField = "injectrol_vi_1"
    .txtinjectrol_srn.DataField = "injectrol_srn"
    
    .txtinjectrol_trh_3.DataField = "injectrol_trh_3"
    .txtinjectrol_trh_2.DataField = "injectrol_trh_2"
    .txtinjectrol_trh_1.DataField = "injectrol_trh_1"
    
    .txtinjectrol_ls_4.DataField = "injectrol_ls_4"
    .txtinjectrol_ls_4d.DataField = "injectrol_ls_4d"
    .txtinjectrol_ls_4c.DataField = "injectrol_ls_4c"
    .txtinjectrol_ls_4b.DataField = "injectrol_ls_4b"
    .txtinjectrol_ls_4a.DataField = "injectrol_ls_4a"
    
    .txtinjectrol_ls_5.DataField = "injectrol_ls_5"
    .txtinjectrol_ls_10.DataField = "injectrol_ls_10"
    .txtinjectrol_back_press.DataField = "injectrol_back_press"
    
    .txtinjectrol_interval.DataField = "injectrol_interval"
    .txtinjectrol_cooling_time.DataField = "injectrol_cooling_time"
    .txtinjectrol_inject_time.DataField = "injectrol_inject_time"
     
    .txtmonitoring_fill.DataField = "monitoring_fill"
    .txtmonitoring_charge.DataField = "monitoring_charge"
    .txtmonitoring_takeout.DataField = "monitoring_takeout"
    .txtmonitoring_cycle.DataField = "monitoring_cycle"
    .txtmonitoring_min_cush.DataField = "monitoring_min_cush"
    .txtmonitoring_act_cush.DataField = "monitoring_act_cush"
    .txtmonitoring_fpc_press.DataField = "monitoring_fpc_press"

    .txtmonitoring_inj_start.DataField = "monitoring_inj_start"
    .txtmonitoring_inj_peak.DataField = "monitoring_inj_peak"
    .txtmonitoring_chg_trq.DataField = "monitoring_chg_trq"
    .txtinjectrol_screw_speed.DataField = "injectrol_screw_speed"

    .txtclamprol_vo_s2.DataField = "clamprol_vo_s2"
    .txtclamprol_vo_3.DataField = "clamprol_vo_3"
    .txtclamprol_vo_2.DataField = "clamprol_vo_2"
    .txtclamprol_vo_1.DataField = "clamprol_vo_1"
    .txtclamprol_vo_s1.DataField = "clamprol_vo_s1"
    
    .txtclamprol_vc_1.DataField = "clamprol_vc_1"
    .txtclamprol_vc_2.DataField = "clamprol_vc_2"
    .txtclamprol_vc_3.DataField = "clamprol_vc_3"
    .txtclamprol_vc_s.DataField = "clamprol_vc_s"
    .txtclamprol_ls_2.DataField = "clamprol_ls_2"
    
    .txtclamprol_pcl.DataField = "clamprol_pcl"
    .txtclamprol_pch.DataField = "clamprol_pch"
    
    .txtclamprol_ls_3.DataField = "clamprol_ls_3"
    .txtclamprol_ls_3b.DataField = "clamprol_ls_3b"
    .txtclamprol_ls_3e.DataField = "clamprol_ls_3e"
    .txtclamprol_ls_3d.DataField = "clamprol_ls_3d"
    .txtclamprol_ls_3a.DataField = "clamprol_ls_3a"

    .txtclamprol_ls_3m.DataField = "clamprol_ls_3m"
'    .txtclamprol_ls_2d.DataField = "clamprol_ls_2d"
'    .txtclamprol_ls_2e.DataField = "clamprol_ls_2e"
'    .txtclamprol_ls_2a.DataField = "clamprol_ls_2a"
'
    .txtclamprol_ve_1.DataField = "clamprol_ve_1"
    .txtclamprol_ve_2.DataField = "clamprol_ve_2"
    .txtclamprol_vr.DataField = "clamprol_vr"
    .txtclamprol_eject_mode.DataField = "clamprol_eject_mode"
    .txtclamprol_eject_count.DataField = "clamprol_eject_count"
    
    .txtclamprol_ls_31a.DataField = "clamprol_ls_31a"
    .txtclamprol_ls_31.DataField = "clamprol_ls_31"
    .txtclamprol_ls_32.DataField = "clamprol_ls_32"
    
    .txtmold_cooling_cavity_mtc.DataField = "mold_cooling_cavity_mtc"
    .txtmold_cooling_cavity_chiller.DataField = "mold_cooling_cavity_chiller"
    .txtmold_cooling_cavity_cooling_twr.DataField = "mold_cooling_cavity_cooling_twr"
    
    .txtmold_cooling_core_mtc.DataField = "mold_cooling_core_mtc"
    .txtmold_cooling_core_chiller.DataField = "mold_cooling_core_chiller"
    .txtmold_cooling_core_cooling_twr.DataField = "mold_cooling_core_cooling_twr"
    
    .txtproduct_data_part_weight.DataField = "product_data_part_weight"
    .txtproduct_data_part_weight_2.DataField = "product_data_part_weight_2"
    .txtproduct_data_runner_weight.DataField = "product_data_runner_weight"
    .txtproduct_data_cavity.DataField = "product_data_cavity"
    
    .txtmold_data_type.DataField = "mold_data_type"
    .txtmold_data_slider.DataField = "mold_data_slider"
    .txtmold_data_core_puller.DataField = "mold_data_core_puller"
    .txtmold_data_air_blow.DataField = "mold_data_air_blow"
    .txtmold_data_hot_rnr_zone.DataField = "mold_data_hot_rnr_zone"
    .txtmold_data_ejector_ls.DataField = "mold_data_ejector_ls"
    .txtmold_data_v.DataField = "mold_data_v"
    .txtmold_data_h.DataField = "mold_data_h"
    .txtmold_data_t.DataField = "mold_data_t"
    .txtmold_data_r_sprue.DataField = "mold_data_r_sprue"
    .txtmold_data_d_sprue.DataField = "mold_data_d_sprue"
    .txtmold_data_hole_sprue_d.DataField = "mold_data_hole_sprue_d"
  
    .txtmaterial_hen.DataField = "material_hen"
    .txtmaterial_hn.DataField = "material_hn"
    .txtmaterial_h1.DataField = "material_hot_runner_1"
    .txtmaterial_h2.DataField = "material_hot_runner_2"
    .txtmaterial_h3.DataField = "material_hot_runner_3"
    .txtmaterial_h4.DataField = "material_hot_runner_4"
    .txtmaterial_hot_runner_1.DataField = "material_h1"
    .txtmaterial_hot_runner_2.DataField = "material_h1"
    .txtmaterial_hot_runner_3.DataField = "material_h1"
    .txtmaterial_hot_runner_4.DataField = "material_h1"
    .txtmaterial_hopper_dehumi.DataField = "material_hopper_dehumi"
    .txtmaterial_drying_time.DataField = "material_drying_time"
    
    .txtmachine_injection_type.DataField = "machine_injection_type"
    .txtmachine_injection_clamping.DataField = "machine_injection_clamping"
    .txtmachine_injection_r_nozzle.DataField = "machine_injection_r_nozzle"
    .txtmachine_injection_d_nozzle.DataField = "machine_injection_d_nozzle"
    .txtmachine_injection_robot_pick_up.DataField = "machine_injection_robot_pick_up"
    .txtmachine_injection_max_shot.DataField = "machine_injection_max_shot"
    .txtmachine_injection_tie_bar_length.DataField = "machine_injection_tie_bar_length"
    .txtmachine_injection_tie_bar_width.DataField = "machine_injection_tie_bar_width"
    
    .txtmachine_injection_nozzle_length.DataField = "machine_injection_nozzle_length"
    .txtmachine_injection_nozzle_diameter.DataField = "machine_injection_nozzle_diameter"
    .txtmachine_injection_core_pack.DataField = "machine_injection_core_pack"
    .txtmachine_injection_min_mould_thickness.DataField = "machine_injection_min_mould_thickness"
    .txtmachine_injection_max_mould_thickness.DataField = "machine_injection_max_mould_thickness"
    .txtmachine_injection_mold_clamping.DataField = "machine_injection_mold_clamping"
End With
End Sub




