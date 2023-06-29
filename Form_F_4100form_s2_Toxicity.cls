VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_F_4100form_s2_Toxicity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Current()
    HighlightNulls
    If CRS = "Yes" Then
        CRSther_anak.Enabled = True
        CRSther_cort.Enabled = True
        CRSther_cortp.Enabled = True
        CRSther_dasa.Enabled = True
        CRSther_etop.Enabled = True
        CRSther_silt.Enabled = True
        CRSther_toci.Enabled = True
        CRSther_othr.Enabled = True
        CRSther_none.Enabled = True
        CRSfev.Enabled = True
        CRShtn.Enabled = True
        CRSmhpx.Enabled = True
        CRSshpx.Enabled = True
        CRSunk.Enabled = True
    Else
        CRSther_anak.Enabled = False
        CRSther_cort.Enabled = False
        CRSther_cortp.Enabled = False
        CRSther_dasa.Enabled = False
        CRSther_etop.Enabled = False
        CRSther_silt.Enabled = False
        CRSther_toci.Enabled = False
        CRSther_othr.Enabled = False
        CRSther_none.Enabled = False
        CRSfev.Enabled = False
        CRShtn.Enabled = False
        CRSmhpx.Enabled = False
        CRSshpx.Enabled = False
        CRSunk.Enabled = False
    End If
    
    CRShtnivf.Enabled = CRShtn
    CRShtnvsp.Enabled = CRShtn
    CRShtnoth.Enabled = CRShtn
    
    CRShtn_vaso_phen.Enabled = CRShtnvsp
    CRShtn_vaso_nore.Enabled = CRShtnvsp
    CRShtn_vaso_epin.Enabled = CRShtnvsp
    CRShtn_vaso_dopa.Enabled = CRShtnvsp
    CRShtn_vaso_vaso.Enabled = CRShtnvsp
    CRShtn_vaso_oth.Enabled = CRShtnvsp
    
    If MASHLH = "Yes" Then
        MASHLH_lab_fib.Enabled = True
        MASHLH_lab_tri.Enabled = True
        MASHLH_lab_non.Enabled = True
    Else
        MASHLH_lab_fib.Enabled = False
        MASHLH_lab_tri.Enabled = False
        MASHLH_lab_non.Enabled = False
    End If
    
        If NTX = "Yes" Then
        NTXdate.Enabled = True
        NTXther_anak.Enabled = True
        NTXther_anti.Enabled = True
        NTXther_cort.Enabled = True
        NTXther_cortp.Enabled = True
        NTXther_silt.Enabled = True
        NTXther_toci.Enabled = True
        NTXther_othr.Enabled = True
        NTXther_none.Enabled = True
        NTXmanif_cered.Enabled = True
        NTXmanif_cva.Enabled = True
        NTXmanif_cogimp.Enabled = True
        NTXmanif_dcons.Enabled = True
        NTXmanif_motor.Enabled = True
        NTXmanif_movdis.Enabled = True
        NTXmanif_perschg.Enabled = True
        NTXmanif_seiz.Enabled = True
        NTXmanif_spch.Enabled = True
        NTXmanif_oth.Enabled = True
    Else
        NTXther_anak.Enabled = False
        NTXther_anti.Enabled = False
        NTXther_cort.Enabled = False
        NTXther_cortp.Enabled = False
        NTXther_silt.Enabled = False
        NTXther_toci.Enabled = False
        NTXther_othr.Enabled = False
        NTXther_none.Enabled = False
        NTXmanif_cered.Enabled = False
        NTXmanif_cva.Enabled = False
        NTXmanif_cogimp.Enabled = False
        NTXmanif_dcons.Enabled = False
        NTXmanif_motor.Enabled = False
        NTXmanif_movdis.Enabled = False
        NTXmanif_perschg.Enabled = False
        NTXmanif_seiz.Enabled = False
        NTXmanif_spch.Enabled = False
        NTXmanif_oth.Enabled = False
    End If
    
    NTXcogimp_amne.Enabled = NTXmanif_cogimp
    NTXcogimp_cogn.Enabled = NTXmanif_cogimp
    NTXcogimp_conc.Enabled = NTXmanif_cogimp
    NTXcogimp_conf.Enabled = NTXmanif_cogimp
    NTXcogimp_ence.Enabled = NTXmanif_cogimp
    NTXcogimp_hall.Enabled = NTXmanif_cogimp
    NTXcogimp_leuk.Enabled = NTXmanif_cogimp
    NTXcogimp_loss.Enabled = NTXmanif_cogimp
    NTXcogimp_ment.Enabled = NTXmanif_cogimp
    NTXcogimp_noni.Enabled = NTXmanif_cogimp
    NTXcogimp_psyc.Enabled = NTXmanif_cogimp
    NTXcogimp_oth.Enabled = NTXmanif_cogimp
    NTXcogimp_othspec.Enabled = NTXmanif_cogimp
    
    NTXmotor_faci.Enabled = NTXmanif_motor
    NTXmotor_guil.Enabled = NTXmanif_motor
    NTXmotor_hemi.Enabled = NTXmanif_motor
    NTXmotor_myel.Enabled = NTXmanif_motor
    NTXmotor_para.Enabled = NTXmanif_motor
    NTXmotor_oth.Enabled = NTXmanif_motor
    NTXmotor_othspec.Enabled = NTXmanif_motor
    
    NTXmovdis_acti.Enabled = NTXmanif_movdis
    NTXmovdis_atax.Enabled = NTXmanif_movdis
    NTXmovdis_cogw.Enabled = NTXmanif_movdis
    NTXmovdis_dysg.Enabled = NTXmanif_movdis
    NTXmovdis_dysk.Enabled = NTXmanif_movdis
    NTXmovdis_dysm.Enabled = NTXmanif_movdis
    NTXmovdis_gait.Enabled = NTXmanif_movdis
    NTXmovdis_myoc.Enabled = NTXmanif_movdis
    NTXmovdis_rest.Enabled = NTXmanif_movdis
    NTXmovdis_oth.Enabled = NTXmanif_movdis
    
    NTXperschg_flat.Enabled = NTXmanif_perschg
    NTXperschg_pers.Enabled = NTXmanif_perschg
    NTXperschg_oth.Enabled = NTXmanif_perschg
End Sub

Private Sub CRS_AfterUpdate()
    If CRS = "Yes" Then
        CRSther_anak.Enabled = True
        CRSther_cort.Enabled = True
        CRSther_cortp.Enabled = True
        CRSther_dasa.Enabled = True
        CRSther_etop.Enabled = True
        CRSther_silt.Enabled = True
        CRSther_toci.Enabled = True
        CRSther_othr.Enabled = True
        CRSther_none.Enabled = True
        CRSfev.Enabled = True
        CRShtn.Enabled = True
        CRSmhpx.Enabled = True
        CRSshpx.Enabled = True
        CRSunk.Enabled = True
    Else
        CRSther_anak.Enabled = False
        CRSther_cort.Enabled = False
        CRSther_cortp.Enabled = False
        CRSther_dasa.Enabled = False
        CRSther_etop.Enabled = False
        CRSther_silt.Enabled = False
        CRSther_toci.Enabled = False
        CRSther_othr.Enabled = False
        CRSther_none.Enabled = False
        CRSfev.Enabled = False
        CRShtn.Enabled = False
        CRSmhpx.Enabled = False
        CRSshpx.Enabled = False
        CRSunk.Enabled = False
    End If
End Sub
Private Sub CRShtn_AfterUpdate()
    CRShtnivf.Enabled = CRShtn
    CRShtnvsp.Enabled = CRShtn
    CRShtnoth.Enabled = CRShtn
End Sub
Private Sub CRShtnvsp_AfterUpdate()
    CRShtn_vaso_phen.Enabled = CRShtnvsp
    CRShtn_vaso_nore.Enabled = CRShtnvsp
    CRShtn_vaso_epin.Enabled = CRShtnvsp
    CRShtn_vaso_dopa.Enabled = CRShtnvsp
    CRShtn_vaso_vaso.Enabled = CRShtnvsp
    CRShtn_vaso_oth.Enabled = CRShtnvsp
End Sub

Private Sub MASHLH_AfterUpdate()
    If MASHLH = "Yes" Then
        MASHLH_lab_fib.Enabled = True
        MASHLH_lab_tri.Enabled = True
        MASHLH_lab_non.Enabled = True
    Else
        MASHLH_lab_fib.Enabled = False
        MASHLH_lab_tri.Enabled = False
        MASHLH_lab_non.Enabled = False
    End If
End Sub

Private Sub NTX_AfterUpdate()
    If NTX = "Yes" Then
        NTXdate.Enabled = True
        NTXther_anak.Enabled = True
        NTXther_anti.Enabled = True
        NTXther_cort.Enabled = True
        NTXther_cortp.Enabled = True
        NTXther_silt.Enabled = True
        NTXther_toci.Enabled = True
        NTXther_othr.Enabled = True
        NTXther_none.Enabled = True
        NTXmanif_cered.Enabled = True
        NTXmanif_cva.Enabled = True
        NTXmanif_cogimp.Enabled = True
        NTXmanif_dcons.Enabled = True
        NTXmanif_motor.Enabled = True
        NTXmanif_movdis.Enabled = True
        NTXmanif_perschg.Enabled = True
        NTXmanif_seiz.Enabled = True
        NTXmanif_spch.Enabled = True
        NTXmanif_oth.Enabled = True
    Else
        NTXther_anak.Enabled = False
        NTXther_anti.Enabled = False
        NTXther_cort.Enabled = False
        NTXther_cortp.Enabled = False
        NTXther_silt.Enabled = False
        NTXther_toci.Enabled = False
        NTXther_othr.Enabled = False
        NTXther_none.Enabled = False
        NTXmanif_cered.Enabled = False
        NTXmanif_cva.Enabled = False
        NTXmanif_cogimp.Enabled = False
        NTXmanif_dcons.Enabled = False
        NTXmanif_motor.Enabled = False
        NTXmanif_movdis.Enabled = False
        NTXmanif_perschg.Enabled = False
        NTXmanif_seiz.Enabled = False
        NTXmanif_spch.Enabled = False
        NTXmanif_oth.Enabled = False
    End If
    HighlightNulls
End Sub
Private Sub NTXassess_AfterUpdate()
    HighlightNulls
End Sub
Private Sub NTXmanif_cogimp_AfterUpdate()
    NTXcogimp_amne.Enabled = NTXmanif_cogimp
    NTXcogimp_cogn.Enabled = NTXmanif_cogimp
    NTXcogimp_conc.Enabled = NTXmanif_cogimp
    NTXcogimp_conf.Enabled = NTXmanif_cogimp
    NTXcogimp_ence.Enabled = NTXmanif_cogimp
    NTXcogimp_hall.Enabled = NTXmanif_cogimp
    NTXcogimp_leuk.Enabled = NTXmanif_cogimp
    NTXcogimp_loss.Enabled = NTXmanif_cogimp
    NTXcogimp_ment.Enabled = NTXmanif_cogimp
    NTXcogimp_noni.Enabled = NTXmanif_cogimp
    NTXcogimp_psyc.Enabled = NTXmanif_cogimp
    NTXcogimp_oth.Enabled = NTXmanif_cogimp
    NTXcogimp_othspec.Enabled = NTXmanif_cogimp
End Sub

Private Sub NTXmanif_motor_AfterUpdate()
    NTXmotor_faci.Enabled = NTXmanif_motor
    NTXmotor_guil.Enabled = NTXmanif_motor
    NTXmotor_hemi.Enabled = NTXmanif_motor
    NTXmotor_myel.Enabled = NTXmanif_motor
    NTXmotor_para.Enabled = NTXmanif_motor
    NTXmotor_oth.Enabled = NTXmanif_motor
    NTXmotor_othspec.Enabled = NTXmanif_motor
End Sub

Private Sub NTXmanif_movdis_AfterUpdate()
    NTXmovdis_acti.Enabled = NTXmanif_movdis
    NTXmovdis_atax.Enabled = NTXmanif_movdis
    NTXmovdis_cogw.Enabled = NTXmanif_movdis
    NTXmovdis_dysg.Enabled = NTXmanif_movdis
    NTXmovdis_dysk.Enabled = NTXmanif_movdis
    NTXmovdis_dysm.Enabled = NTXmanif_movdis
    NTXmovdis_gait.Enabled = NTXmanif_movdis
    NTXmovdis_myoc.Enabled = NTXmanif_movdis
    NTXmovdis_rest.Enabled = NTXmanif_movdis
    NTXmovdis_oth.Enabled = NTXmanif_movdis
End Sub

Private Sub NTXmanif_perschg_AfterUpdate()
    NTXperschg_flat.Enabled = NTXmanif_perschg
    NTXperschg_pers.Enabled = NTXmanif_perschg
    NTXperschg_oth.Enabled = NTXmanif_perschg
End Sub

'Private Sub gd3tox_AfterUpdate()
'    If gd3tox = "Yes" Then
'        gd3tox_org1.Enabled = True
'        gd3tox_org2.Enabled = True
'        gd3tox_org3.Enabled = True
'        gd3tox_tox1.Enabled = True
'        gd3tox_tox2.Enabled = True
'        gd3tox_tox3.Enabled = True
'        gd3tox_dt1.Enabled = True
'        gd3tox_dt2.Enabled = True
'        gd3tox_dt3.Enabled = True
'        gd3tox_res1.Enabled = True
'        gd3tox_res2.Enabled = True
'        gd3tox_res3.Enabled = True
'        gd3tox_res_dt1.Enabled = True
'        gd3tox_res_dt2.Enabled = True
'        gd3tox_res_dt3.Enabled = True
'    Else
'        gd3tox_org1.Enabled = False
'        gd3tox_org2.Enabled = False
'        gd3tox_org3.Enabled = False
'        gd3tox_tox1.Enabled = False
'        gd3tox_tox2.Enabled = False
'        gd3tox_tox3.Enabled = False
'        gd3tox_dt1.Enabled = False
'        gd3tox_dt2.Enabled = False
'        gd3tox_dt3.Enabled = False
'        gd3tox_res1.Enabled = False
'        gd3tox_res2.Enabled = False
'        gd3tox_res3.Enabled = False
'        gd3tox_res_dt1.Enabled = False
'        gd3tox_res_dt2.Enabled = False
'        gd3tox_res_dt3.Enabled = False
'    End If
'End Sub
'
'Private Sub gr4tox_AfterUpdate()
'    If gd4tox = "Yes" Then
'        gd4tox_org1.Enabled = True
'        gd4tox_org2.Enabled = True
'        gd4tox_org3.Enabled = True
'        gd4tox_tox1.Enabled = True
'        gd4tox_tox2.Enabled = True
'        gd4tox_tox3.Enabled = True
'        gd4tox_dt1.Enabled = True
'        gd4tox_dt2.Enabled = True
'        gd4tox_dt3.Enabled = True
'        gd4tox_res1.Enabled = True
'        gd4tox_res2.Enabled = True
'        gd4tox_res3.Enabled = True
'        gd4tox_res_dt1.Enabled = True
'        gd4tox_res_dt2.Enabled = True
'        gd4tox_res_dt3.Enabled = True
'    Else
'        gd4tox_org1.Enabled = False
'        gd4tox_org2.Enabled = False
'        gd4tox_org3.Enabled = False
'        gd4tox_tox1.Enabled = False
'        gd4tox_tox2.Enabled = False
'        gd4tox_tox3.Enabled = False
'        gd4tox_dt1.Enabled = False
'        gd4tox_dt2.Enabled = False
'        gd4tox_dt3.Enabled = False
'        gd4tox_res1.Enabled = False
'        gd4tox_res2.Enabled = False
'        gd4tox_res3.Enabled = False
'        gd4tox_res_dt1.Enabled = False
'        gd4tox_res_dt2.Enabled = False
'        gd4tox_res_dt3.Enabled = False
'    End If
'End Sub

Private Sub HighlightNulls()
    If IsNull(NTXassess) Then
        NTXassess.BackColor = RGB(255, 242, 0)
    Else
        NTXassess.BackColor = RGB(255, 255, 255)
    End If
End Sub
