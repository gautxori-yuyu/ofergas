Attribute VB_Name = "modOfertaTypes"
'=========================================================
' M�DULO DE TIPOS PARA OFERTAS Y CAT�LOGOS
' Autor: Generado para sistema ABC_ofertas
' Fecha: Diciembre 2024
'=========================================================
Option Explicit

'=========================================================
' TIPOS PARA DATOS GENERALES DE OFERTA
'=========================================================

Public Type tOfertasDatosGenerales
    OFER_ID As String
    OFER_NUM_OFERTA As String
    OFER_FECHA As Date
    OFER_CLIENTE As String
    OFER_USUARIO_FINAL As String
    pais_id As Long
    COME_ID As Long
    OFER_NUM_CALCULO As String
    GASE_ID As Long
    OFER_OBSERVACIONES As String
    OFER_ESTADO_OFERTA As Long
    OFTN_ID As Long
    OFER_FEC_ULT_MODIF As Date
    OFER_CALD_EXTRA As Double
    OFER_CALD_RADIO As Long
    OFER_CALD_ASME As Double
    OFER_CALD_INOX As Double
    OFER_CALD_SELLO As Long
    ' Campos resueltos de cat�logos
    pais_nombre As String
    COME_NOMBRE As String
    OFTN_NOMBRE As String
    GASE_DENOMINACION As String
End Type

'=========================================================
' TIPOS PARA TABLAS HIJAS DE OFERTA
'=========================================================

Public Type tOfertasManoObra
    OFMA_ID As String
    OFER_ID As String
    MOCA_ID As Double
    OFMA_MOCA_FASE1_HORAS As Double
    OFMA_MOCA_FASE1_PRECIO As Double
    OFMA_MOCA_FASE2_HORAS As Double
    OFMA_MOCA_FASE2_PRECIO As Double
    OFMA_MOCA_SOLDA_HORAS As Double
    OFMA_MOCA_SOLDA_PRECIO As Double
    OFMA_MOCA_PROBA_HORAS As Double
    OFMA_MOCA_PROBA_PRECIO As Double
    OFMA_MOCA_PINTU_HORAS As Double
    OFMA_MOCA_PINTU_PRECIO As Double
    OFMA_MOCA_ELECT_HORAS As Double
    OFMA_MOCA_ELECT_PRECIO As Double
    MOIN_ID As Double
    OFMA_MOIN_INGEN_HORAS As Double
    OFMA_MOIN_INGEN_PRECIO As Double
    OFMA_IMPORTE As Double
    ' Flags
    ImporteCalculado As Boolean
End Type

Public Type tOfertasModelo
    OFMO_ID As String
    OFER_ID As String
    ' Chasis
    CHAS_ID As Double
    OFMO_CHAS_PRE_COSTE As Double
    OFMO_CHAS_IMPORTE As Double
    OFMO_CHAS_DESCRIPCION As String
    OFMO_CHAS_OBSERVA As String
    OFMO_CHAS_ADVERTENCIA As String
    ' Correas
    CORR_ID As Double
    OFMO_CORR_PRE_COSTE As Double
    OFMO_CORR_IMPORTE As Double
    OFMO_CORR_DESCRIPCION As String
    OFMO_CORR_OBSERVA As String
    OFMO_CORR_ADVERTENCIA As String
    ' Acoplamientos directos
    ACOD_ID As Double
    OFMO_ACOD_PRE_COSTE As Double
    OFMO_ACOD_IMPORTE As Double
    OFMO_ACOD_DESCRIPCION As String
    OFMO_ACOD_OBSERVA As String
    OFMO_ACOD_ADVERTENCIA As String
    ' Reductores
    REDU_ID As Double
    OFMO_REDU_PRE_COSTE As Double
    OFMO_REDU_IMPORTE As Double
    OFMO_REDU_DESCRIPCION As String
    OFMO_REDU_OBSERVA As String
    OFMO_REDU_ADVERTENCIA As String
    ' Transmisi�n
    OFMO_TRAN_TIPO As Long
    OFMO_ATEX_SN As String
    OFMO_ATEX_PRE_COSTE As Double
    ' Tuber�as aire
    TUAI_ID As Long
    TUDI_ID As Double
    OFMO_TUDI_PRE_COSTE As Double
    OFMO_TUDI_IMPORTE As Double
    OFMO_TUDI_DESCRIPCION As String
    OFMO_TUDI_OBSERVA As String
    OFMO_TUDI_ADVERTENCIA As String
    OFMO_TUDI_CANTIDAD As Long
    ' Tuber�as agua
    TUAG_ID As Double
    OFMO_TUAG_PRE_COSTE As Double
    OFMO_TUAG_IMPORTE As Double
    OFMO_TUAG_DESCRIPCION As String
    OFMO_TUAG_OBSERVA As String
    OFMO_TUAG_ADVERTENCIA As String
    ' Flags
    ImporteCalculado As Boolean
End Type

Public Type tOfertasOpciones
    OFOP_ID As String
    OFER_ID As String
    ' V�lvula retenci�n
    VARE_ID As Double
    OFOP_VARE_PRE_COSTE As Double
    OFOP_VARE_IMPORTE As Double
    OFOP_VARE_DESCRIPCION As String
    OFOP_VARE_OBSERVA As String
    OFOP_VARE_ADVERTENCIA As String
    ' Electrov�lvula
    ELEC_ID As Double
    OFOP_ELEC_PRE_COSTE As Double
    OFOP_ELEC_IMPORTE As Double
    OFOP_ELEC_DESCRIPCION As String
    OFOP_ELEC_OBSERVA As String
    OFOP_ELEC_ADVERTENCIA As String
    ' V�lvula presi�n
    VAPR_ID As Double
    OFOP_VAPR_PRE_COSTE As Double
    OFOP_VAPR_IMPORTE As Double
    OFOP_VAPR_DESCRIPCION As String
    OFOP_VAPR_OBSERVA As String
    OFOP_VAPR_ADVERTENCIA As String
    ' Engrase cilindros
    ENCI_ID As Double
    OFOP_ENCI_PRE_COSTE As Double
    OFOP_ENCI_IMPORTE As Double
    OFOP_ENCI_DESCRIPCION As String
    OFOP_ENCI_OBSERVA As String
    OFOP_ENCI_ADVERTENCIA As String
    ' Purgador
    PURG_ID As Double
    OFOP_PURG_PRE_COSTE As Double
    OFOP_PURG_IMPORTE As Double
    OFOP_PURG_DESCRIPCION As String
    OFOP_PURG_OBSERVA As String
    OFOP_PURG_ADVERTENCIA As String
    OFOP_PURG_CANTIDAD As Integer
    ' Resistencia calefacci�n
    RESC_ID As Double
    OFOP_RESC_PRE_COSTE As Double
    OFOP_RESC_IMPORTE As Double
    OFOP_RESC_DESCRIPCION As String
    OFOP_RESC_OBSERVA As String
    OFOP_RESC_ADVERTENCIA As String
    ' V�lvula reguladora
    VARG_ID As Double
    OFOP_VARG_PRE_COSTE As Double
    OFOP_VARG_IMPORTE As Double
    OFOP_VARG_DESCRIPCION As String
    OFOP_VARG_OBSERVA As String
    OFOP_VARG_ADVERTENCIA As String
    ' Flags
    ImporteCalculado As Boolean
End Type

Public Type tOfertasCabezal
    OFCA_ID As String
    OFER_ID As String
    ' Cabezal
    CABE_ID As Double
    OFCA_CABE_PRE_COSTE As Double
    OFCA_CABE_IMPORTE As Double
    OFCA_CABE_DESCRIPCION As String
    OFCA_CABE_OBSERVA As String
    OFCA_CABE_ADVERTENCIA As String
    OFCA_NUM_ETAPAS As Long
    ' Cilindro 1
    CIL1_ID As Long
    OFCA_CIL1_PRE_COSTE As Double
    OFCA_CIL1_IMPORTE As Double
    OFCA_CIL1_DESCRIPCION As String
    OFCA_CIL1_OBSERVA As String
    OFCA_CIL1_ADVERTENCIA As String
    OFCA_CIL1_CANTIDAD As Long
    OFCA_CIL1_EMPA_SN As String
    OFCA_CIL1_EMPA_PRE_COSTE As Double
    OFCA_CIL1_ANSE_PRE_COSTE As Double
    OFCA_CIL1_BSAS_PRE_COSTE As Double
    ' Cilindro 2
    CIL2_ID As Long
    OFCA_CIL2_PRE_COSTE As Double
    OFCA_CIL2_IMPORTE As Double
    OFCA_CIL2_DESCRIPCION As String
    OFCA_CIL2_OBSERVA As String
    OFCA_CIL2_ADVERTENCIA As String
    OFCA_CIL2_CANTIDAD As Long
    OFCA_CIL2_EMPA_SN As String
    OFCA_CIL2_EMPA_PRE_COSTE As Double
    OFCA_CIL2_ANSE_PRE_COSTE As Double
    OFCA_CIL2_BSAS_PRE_COSTE As Double
    ' Cilindro 3
    CIL3_ID As Long
    OFCA_CIL3_PRE_COSTE As Double
    OFCA_CIL3_IMPORTE As Double
    OFCA_CIL3_DESCRIPCION As String
    OFCA_CIL3_OBSERVA As String
    OFCA_CIL3_ADVERTENCIA As String
    OFCA_CIL3_CANTIDAD As Long
    OFCA_CIL3_EMPA_SN As String
    OFCA_CIL3_EMPA_PRE_COSTE As Double
    OFCA_CIL3_ANSE_PRE_COSTE As Double
    OFCA_CIL3_BSAS_PRE_COSTE As Double
    ' Cilindro 4
    CIL4_ID As Long
    OFCA_CIL4_PRE_COSTE As Double
    OFCA_CIL4_IMPORTE As Double
    OFCA_CIL4_DESCRIPCION As String
    OFCA_CIL4_OBSERVA As String
    OFCA_CIL4_ADVERTENCIA As String
    OFCA_CIL4_CANTIDAD As Long
    OFCA_CIL4_EMPA_SN As String
    OFCA_CIL4_EMPA_PRE_COSTE As Double
    OFCA_CIL4_ANSE_PRE_COSTE As Double
    OFCA_CIL4_BSAS_PRE_COSTE As Double
    ' Cilindro 5
    CIL5_ID As Long
    OFCA_CIL5_PRE_COSTE As Double
    OFCA_CIL5_IMPORTE As Double
    OFCA_CIL5_DESCRIPCION As String
    OFCA_CIL5_OBSERVA As String
    OFCA_CIL5_ADVERTENCIA As String
    OFCA_CIL5_CANTIDAD As Long
    OFCA_CIL5_EMPA_SN As String
    OFCA_CIL5_EMPA_PRE_COSTE As Double
    OFCA_CIL5_ANSE_PRE_COSTE As Double
    OFCA_CIL5_BSAS_PRE_COSTE As Double
    ' Cilindro 6
    CIL6_ID As Long
    OFCA_CIL6_PRE_COSTE As Double
    OFCA_CIL6_IMPORTE As Double
    OFCA_CIL6_DESCRIPCION As String
    OFCA_CIL6_OBSERVA As String
    OFCA_CIL6_ADVERTENCIA As String
    OFCA_CIL6_CANTIDAD As Long
    OFCA_CIL6_EMPA_SN As String
    OFCA_CIL6_EMPA_PRE_COSTE As Double
    OFCA_CIL6_ANSE_PRE_COSTE As Double
    OFCA_CIL6_BSAS_PRE_COSTE As Double
    ' Anti-surge
    OFCA_ANSE_TIPO As Long
    ' Bloque SAS
    OFCA_BSAS_SN As String
    OFCA_BSAS_PRE_COSTE As Double
    OFCA_BSAS_IMPORTE As Double
    ' Normativa compresor
    NORC_ID As Long
    OFCA_NORC_PRE_COSTE As Double
    OFCA_NORC_IMPORTE As Double
    OFCA_NORC_DESCRIPCION As String
    OFCA_NORC_OBSERVA As String
    OFCA_NORC_ADVERTENCIA As String
    ' Flags
    ImporteCalculado As Boolean
End Type

Public Type tOfertasCalderines
    OFCL_ID As Long
    OFER_ID As String
    ' Calder�n base
    CALD_ID As Long
    ' Etapa 1 - Volumen aire
    CVA1_ID As Long
    CLA1_ID As Long
    OFCL_CLA1_MATERIAL As Long
    OFCL_CLA1_PRE_COSTE As Double
    OFCL_CLA1_DESCRIPCION As String
    OFCL_CLA1_OBSERVA As String
    OFCL_CLA1_ADVERTENCIA As String
    OFCL_CLA1_CANTIDAD As Long
    ' Etapa 1 - Volumen entrada
    CVE1_ID As Long
    CLE1_ID As Long
    OFCL_CLE1_MATERIAL As Long
    OFCL_CLE1_PRE_COSTE As Double
    OFCL_CLE1_DESCRIPCION As String
    OFCL_CLE1_OBSERVA As String
    OFCL_CLE1_ADVERTENCIA As String
    OFCL_CLE1_CANTIDAD As Long
    OFCL_ETP1_IMPORTE As Double
    ' Etapa 2 - Volumen aire
    CVA2_ID As Long
    CLA2_ID As Long
    OFCL_CLA2_MATERIAL As Long
    OFCL_CLA2_PRE_COSTE As Double
    OFCL_CLA2_DESCRIPCION As String
    OFCL_CLA2_OBSERVA As String
    OFCL_CLA2_ADVERTENCIA As String
    OFCL_CLA2_CANTIDAD As Long
    ' Etapa 2 - Volumen entrada
    CVE2_ID As Long
    CLE2_ID As Long
    OFCL_CLE2_MATERIAL As Long
    OFCL_CLE2_PRE_COSTE As Double
    OFCL_CLE2_DESCRIPCION As String
    OFCL_CLE2_OBSERVA As String
    OFCL_CLE2_ADVERTENCIA As String
    OFCL_CLE2_CANTIDAD As Long
    OFCL_ETP2_IMPORTE As Double
    ' Etapa 3 - Volumen aire
    CVA3_ID As Long
    CLA3_ID As Long
    OFCL_CLA3_MATERIAL As Long
    OFCL_CLA3_PRE_COSTE As Double
    OFCL_CLA3_DESCRIPCION As String
    OFCL_CLA3_OBSERVA As String
    OFCL_CLA3_ADVERTENCIA As String
    OFCL_CLA3_CANTIDAD As Long
    ' Etapa 3 - Volumen entrada
    CVE3_ID As Long
    CLE3_ID As Long
    OFCL_CLE3_MATERIAL As Long
    OFCL_CLE3_PRE_COSTE As Double
    OFCL_CLE3_DESCRIPCION As String
    OFCL_CLE3_OBSERVA As String
    OFCL_CLE3_ADVERTENCIA As String
    OFCL_CLE3_CANTIDAD As Long
    OFCL_ETP3_IMPORTE As Double
    ' Etapa 4 - Volumen aire
    CVA4_ID As Long
    CLA4_ID As Long
    OFCL_CLA4_MATERIAL As Long
    OFCL_CLA4_PRE_COSTE As Double
    OFCL_CLA4_DESCRIPCION As String
    OFCL_CLA4_OBSERVA As String
    OFCL_CLA4_ADVERTENCIA As String
    OFCL_CLA4_CANTIDAD As Long
    ' Etapa 4 - Volumen entrada
    CVE4_ID As Long
    CLE4_ID As Long
    OFCL_CLE4_MATERIAL As Long
    OFCL_CLE4_PRE_COSTE As Double
    OFCL_CLE4_DESCRIPCION As String
    OFCL_CLE4_OBSERVA As String
    OFCL_CLE4_ADVERTENCIA As String
    OFCL_CLE4_CANTIDAD As Long
    OFCL_ETP4_IMPORTE As Double
    ' Etapa 5 - Volumen aire
    CVA5_ID As Long
    CLA5_ID As Long
    OFCL_CLA5_MATERIAL As Long
    OFCL_CLA5_PRE_COSTE As Double
    OFCL_CLA5_DESCRIPCION As String
    OFCL_CLA5_OBSERVA As String
    OFCL_CLA5_ADVERTENCIA As String
    OFCL_CLA5_CANTIDAD As Long
    ' Etapa 5 - Volumen entrada
    CVE5_ID As Long
    CLE5_ID As Long
    OFCL_CLE5_MATERIAL As Long
    OFCL_CLE5_PRE_COSTE As Double
    OFCL_CLE5_DESCRIPCION As String
    OFCL_CLE5_OBSERVA As String
    OFCL_CLE5_ADVERTENCIA As String
    OFCL_CLE5_CANTIDAD As Long
    OFCL_ETP5_IMPORTE As Double
    ' Etapa 6 - Volumen aire
    CVA6_ID As Long
    CLA6_ID As Long
    OFCL_CLA6_MATERIAL As Long
    OFCL_CLA6_PRE_COSTE As Double
    OFCL_CLA6_DESCRIPCION As String
    OFCL_CLA6_OBSERVA As String
    OFCL_CLA6_ADVERTENCIA As String
    OFCL_CLA6_CANTIDAD As Long
    ' Etapa 6 - Volumen entrada
    CVE6_ID As Long
    CLE6_ID As Long
    OFCL_CLE6_MATERIAL As Long
    OFCL_CLE6_PRE_COSTE As Double
    OFCL_CLE6_DESCRIPCION As String
    OFCL_CLE6_OBSERVA As String
    OFCL_CLE6_ADVERTENCIA As String
    OFCL_CLE6_CANTIDAD As Long
    OFCL_ETP6_IMPORTE As Double
    ' Opciones
    OFCL_EXTRA_SN As String
    OFCL_RADIO_SN As String
    OFCL_ASME_SN As String
    OFCL_SELLOU_SN As String
    OFCL_SELLOU_IMPORTE As Long
    ' Dep�sitos entrada
    DEPE_ID As Long
    OFCL_DEPE_PRE_COSTE As Double
    OFCL_DEPE_IMPORTE As Double
    OFCL_DEPE_DESCRIPCION As String
    OFCL_DEPE_OBSERVA As String
    OFCL_DEPE_ADVERTENCIA As String
    ' Dep�sitos final
    DEPF_ID As Long
    OFCL_DEPF_PRE_COSTE As Double
    OFCL_DEPF_IMPORTE As Double
    OFCL_DEPF_DESCRIPCION As String
    OFCL_DEPF_OBSERVA As String
    OFCL_DEPF_ADVERTENCIA As String
    ' Flags
    ImporteCalculado As Boolean
End Type

Public Type tOfertasRefrigeradores
    OFRF_ID As String
    OFER_ID As String
    ' Tipo tubos
    TITU_ID As Long
    ' Refrigerador 1
    REF1_ID As Long
    OFRF_REF1_TIPO As Long
    OFRF_REF1_PN As Long
    OFRF_REF1_CE As Long
    OFRF_REF1_CS As Long
    OFRF_REF1_PRE_COSTE As Double
    OFRF_REF1_CE_COSTE As Double
    OFRF_REF1_CS_COSTE As Double
    OFRF_REF1_IMPORTE As Double
    OFRF_REF1_DESCRIPCION As String
    OFRF_REF1_OBSERVA As String
    OFRF_REF1_ADVERTENCIA As String
    OFRF_REF1_CANTIDAD As Long
    ' Refrigerador 2
    REF2_ID As Long
    OFRF_REF2_TIPO As Long
    OFRF_REF2_PN As Long
    OFRF_REF2_CE As Long
    OFRF_REF2_CS As Long
    OFRF_REF2_PRE_COSTE As Double
    OFRF_REF2_CE_COSTE As Double
    OFRF_REF2_CS_COSTE As Double
    OFRF_REF2_IMPORTE As Double
    OFRF_REF2_DESCRIPCION As String
    OFRF_REF2_OBSERVA As String
    OFRF_REF2_ADVERTENCIA As String
    OFRF_REF2_CANTIDAD As Long
    ' Refrigerador 3
    REF3_ID As Long
    OFRF_REF3_TIPO As Long
    OFRF_REF3_PN As Long
    OFRF_REF3_CE As Long
    OFRF_REF3_CS As Long
    OFRF_REF3_PRE_COSTE As Double
    OFRF_REF3_CE_COSTE As Double
    OFRF_REF3_CS_COSTE As Double
    OFRF_REF3_IMPORTE As Double
    OFRF_REF3_DESCRIPCION As String
    OFRF_REF3_OBSERVA As String
    OFRF_REF3_ADVERTENCIA As String
    OFRF_REF3_CANTIDAD As Long
    ' Refrigerador 4
    REF4_ID As Long
    OFRF_REF4_TIPO As Long
    OFRF_REF4_PN As Long
    OFRF_REF4_CE As Long
    OFRF_REF4_CS As Long
    OFRF_REF4_PRE_COSTE As Double
    OFRF_REF4_CE_COSTE As Double
    OFRF_REF4_CS_COSTE As Double
    OFRF_REF4_IMPORTE As Double
    OFRF_REF4_DESCRIPCION As String
    OFRF_REF4_OBSERVA As String
    OFRF_REF4_ADVERTENCIA As String
    OFRF_REF4_CANTIDAD As Long
    ' Refrigerador 5
    REF5_ID As Long
    OFRF_REF5_TIPO As Long
    OFRF_REF5_PN As Long
    OFRF_REF5_CE As Long
    OFRF_REF5_CS As Long
    OFRF_REF5_PRE_COSTE As Double
    OFRF_REF5_CE_COSTE As Double
    OFRF_REF5_CS_COSTE As Double
    OFRF_REF5_IMPORTE As Double
    OFRF_REF5_DESCRIPCION As String
    OFRF_REF5_OBSERVA As String
    OFRF_REF5_ADVERTENCIA As String
    OFRF_REF5_CANTIDAD As Long
    ' Refrigerador 6
    REF6_ID As Long
    OFRF_REF6_TIPO As Long
    OFRF_REF6_PN As Long
    OFRF_REF6_CE As Long
    OFRF_REF6_CS As Long
    OFRF_REF6_PRE_COSTE As Double
    OFRF_REF6_CE_COSTE As Double
    OFRF_REF6_CS_COSTE As Double
    OFRF_REF6_IMPORTE As Double
    OFRF_REF6_DESCRIPCION As String
    OFRF_REF6_OBSERVA As String
    OFRF_REF6_ADVERTENCIA As String
    OFRF_REF6_CANTIDAD As Long
    ' Opciones
    OFRF_EXTRA_SN As String
    OFRF_RADIO_SN As String
    OFRF_ASME_SN As String
    OFRF_SELLOU_SN As String
    OFRF_SELLOU_IMPORTE As Long
    ' V�lvula seguridad
    VASE_ID As Long
    OFRF_VASE_PRE_COSTE As Double
    OFRF_VASE_IMPORTE As Double
    OFRF_VASE_DESCRIPCION As String
    OFRF_VASE_OBSERVA As String
    OFRF_VASE_ADVERTENCIA As String
    OFRF_VASE_CANTIDAD As Long
    ' Flags
    ImporteCalculado As Boolean
End Type

Public Type tOfertasAccesorios
    OFAC_ID As String
    OFER_ID As String
    ' Motor
    MOTO_ID As Long
    OFAC_MOTO_PRE_COSTE As Double
    OFAC_MOTO_IMPORTE As Double
    OFAC_MOTO_DESCRIPCION As String
    OFAC_MOTO_OBSERVA As String
    OFAC_MOTO_ADVERTENCIA As String
    OFAC_MOTO_TIPO As Long
    OFAC_MOTO_ESP As String
    OFAC_MOTO_ESP_IMPORTE As Double
    ' Arrancador (l�gica especial seg�n TIPO)
    ARRA_ID As Long
    OFAC_ARRA_PRE_COSTE As Double
    OFAC_ARRA_IMPORTE As Double
    OFAC_ARRA_DESCRIPCION As String
    OFAC_ARRA_OBSERVA As String
    OFAC_ARRA_ADVERTENCIA As String
    OFAC_ARRA_CANTIDAD As Long
    OFAC_ARRA_TIPO As Long
    ' Caja local
    CALO_ID As Long
    OFAC_CALO_PRE_COSTE As Double
    OFAC_CALO_IMPORTE As Double
    OFAC_CALO_DESCRIPCION As String
    OFAC_CALO_OBSERVA As String
    OFAC_CALO_ADVERTENCIA As String
    ' Filtro
    FILT_ID As Long
    OFAC_FILT_PRE_COSTE As Double
    OFAC_FILT_IMPORTE As Double
    OFAC_FILT_DESCRIPCION As String
    OFAC_FILT_OBSERVA As String
    OFAC_FILT_ADVERTENCIA As String
    ' Aero
    AERO_ID As Double
    OFAC_AERO_PRE_COSTE As Double
    OFAC_AERO_IMPORTE As Double
    OFAC_AERO_DESCRIPCION As String
    OFAC_AERO_OBSERVA As String
    OFAC_AERO_ADVERTENCIA As String
    ' Llave entrada
    LLEN_ID As Double
    OFAC_LLEN_PRE_COSTE As Double
    OFAC_LLEN_IMPORTE As Double
    OFAC_LLEN_DESCRIPCION As String
    OFAC_LLEN_OBSERVA As String
    OFAC_LLEN_ADVERTENCIA As String
    ' Llave salida
    LLSA_ID As Double
    OFAC_LLSA_PRE_COSTE As Double
    OFAC_LLSA_IMPORTE As Double
    OFAC_LLSA_DESCRIPCION As String
    OFAC_LLSA_OBSERVA As String
    OFAC_LLSA_ADVERTENCIA As String
    ' Grupo engrase
    GREN_ID As Double
    OFAC_GREN_PRE_COSTE As Double
    OFAC_GREN_IMPORTE As Double
    OFAC_GREN_DESCRIPCION As String
    OFAC_GREN_OBSERVA As String
    OFAC_GREN_ADVERTENCIA As String
    ' Flags
    ImporteCalculado As Boolean
End Type

Public Type tOfertasInstrumentacion
    OFIN_ID As String
    OFER_ID As String
    ' Base instrumentaci�n
    INTR_ID As Double
    OFIN_INTR_DESCRIPCION As String
    OFIN_INTR_OBSERVA As String
    OFIN_INTR_ADVERTENCIA As String
    ' Transmisor temperatura
    TRAT_ID As Double
    OFIN_TRAT_PRE_COSTE As Double
    OFIN_TRAT_IMPORTE As Double
    OFIN_TRAT_DESCRIPCION As String
    OFIN_TRAT_OBSERVA As String
    OFIN_TRAT_ADVERTENCIA As String
    OFIN_TRAT_CANTIDAD As Long
    ' Transmisor presi�n
    TRAP_ID As Double
    OFIN_TRAP_PRE_COSTE As Double
    OFIN_TRAP_IMPORTE As Double
    OFIN_TRAP_DESCRIPCION As String
    OFIN_TRAP_OBSERVA As String
    OFIN_TRAP_ADVERTENCIA As String
    OFIN_TRAP_CANTIDAD As Long
    ' Term�metro
    TERM_ID As Double
    OFIN_TERM_PRE_COSTE As Double
    OFIN_TERM_IMPORTE As Double
    OFIN_TERM_DESCRIPCION As String
    OFIN_TERM_OBSERVA As String
    OFIN_TERM_ADVERTENCIA As String
    OFIN_TERM_CANTIDAD As Long
    ' Man�metro
    MANO_ID As Double
    OFIN_MANO_PRE_COSTE As Double
    OFIN_MANO_IMPORTE As Double
    OFIN_MANO_DESCRIPCION As String
    OFIN_MANO_OBSERVA As String
    OFIN_MANO_ADVERTENCIA As String
    OFIN_MANO_CANTIDAD As Long
    ' Electrov�lvula regulaci�n
    ELER_ID As Double
    OFIN_ELER_PRE_COSTE As Double
    OFIN_ELER_IMPORTE As Double
    OFIN_ELER_DESCRIPCION As String
    OFIN_ELER_OBSERVA As String
    OFIN_ELER_ADVERTENCIA As String
    OFIN_ELER_CANTIDAD As Long
    ' Sensor ca�da v�stago
    SECV_ID As Double
    OFIN_SECV_PRE_COSTE As Double
    OFIN_SECV_IMPORTE As Double
    OFIN_SECV_DESCRIPCION As String
    OFIN_SECV_OBSERVA As String
    OFIN_SECV_ADVERTENCIA As String
    OFIN_SECV_CANTIDAD As Long
    ' Interruptor vibraci�n
    INTV_ID As Double
    OFIN_INTV_PRE_COSTE As Double
    OFIN_INTV_IMPORTE As Double
    OFIN_INTV_DESCRIPCION As String
    OFIN_INTV_OBSERVA As String
    OFIN_INTV_ADVERTENCIA As String
    OFIN_INTV_CANTIDAD As Long
    ' Interruptor nivel aceite
    INTA_ID As Double
    OFIN_INTA_PRE_COSTE As Double
    OFIN_INTA_IMPORTE As Double
    OFIN_INTA_DESCRIPCION As String
    OFIN_INTA_OBSERVA As String
    OFIN_INTA_ADVERTENCIA As String
    OFIN_INTA_CANTIDAD As Long
    ' Nivel condensados
    NIVC_ID As Double
    OFIN_NIVC_PRE_COSTE As Double
    OFIN_NIVC_IMPORTE As Double
    OFIN_NIVC_DESCRIPCION As String
    OFIN_NIVC_OBSERVA As String
    OFIN_NIVC_ADVERTENCIA As String
    OFIN_NIVC_CANTIDAD As Long
    ' V�lvula termost�tica
    VATE_ID As Double
    OFIN_VATE_PRE_COSTE As Double
    OFIN_VATE_IMPORTE As Double
    OFIN_VATE_DESCRIPCION As String
    OFIN_VATE_OBSERVA As String
    OFIN_VATE_ADVERTENCIA As String
    OFIN_VATE_CANTIDAD As Long
    ' Instrumentaci�n libre (7 l�neas)
    OFIN_INS1_TEXTO As String
    OFIN_INS1_CANTIDAD As Integer
    OFIN_INS1_IMPORTE As Double
    OFIN_INS2_TEXTO As String
    OFIN_INS2_CANTIDAD As Integer
    OFIN_INS2_IMPORTE As Double
    OFIN_INS3_TEXTO As String
    OFIN_INS3_CANTIDAD As Integer
    OFIN_INS3_IMPORTE As Double
    OFIN_INS4_TEXTO As String
    OFIN_INS4_CANTIDAD As Integer
    OFIN_INS4_IMPORTE As Double
    OFIN_INS5_TEXTO As String
    OFIN_INS5_CANTIDAD As Integer
    OFIN_INS5_IMPORTE As Double
    OFIN_INS6_TEXTO As String
    OFIN_INS6_CANTIDAD As Integer
    OFIN_INS6_IMPORTE As Double
    OFIN_INS7_TEXTO As String
    OFIN_INS7_CANTIDAD As Integer
    OFIN_INS7_IMPORTE As Double
    ' Flags
    ImporteCalculado As Boolean
End Type

Public Type tOfertasExtras
    OFEX_ID As String
    OFER_ID As String
    ' Embalaje
    EMBA_ID As Long
    OFEX_EMBA_PRE_COSTE As Double
    OFEX_EMBA_IMPORTE As Double
    OFEX_EMBA_DESCRIPCION As String
    OFEX_EMBA_OBSERVA As String
    OFEX_EMBA_ADVERTENCIA As String
    ' Transporte
    TRAN_ID As Long
    OFEX_TRAN_PRE_COSTE As Double
    OFEX_TRAN_IMPORTE As Double
    OFEX_TRAN_DESCRIPCION As String
    OFEX_TRAN_OBSERVA As String
    OFEX_TRAN_ADVERTENCIA As String
    ' Puesta en marcha
    PUMA_ID As Long
    OFEX_PUMA_PRE_COSTE As Double
    OFEX_PUMA_IMPORTE As Double
    OFEX_PUMA_DESCRIPCION As String
    OFEX_PUMA_OBSERVA As String
    OFEX_PUMA_ADVERTENCIA As String
    ' Pruebas asistencia 1
    PRA1_ID As Long
    OFEX_PRA1_PRE_COSTE As Double
    OFEX_PRA1_IMPORTE As Double
    OFEX_PRA1_DESCRIPCION As String
    OFEX_PRA1_OBSERVA As String
    OFEX_PRA1_ADVERTENCIA As String
    ' Pruebas asistencia 2
    PRA2_ID As Long
    OFEX_PRA2_PRE_COSTE As Double
    OFEX_PRA2_IMPORTE As Double
    OFEX_PRA2_DESCRIPCION As String
    OFEX_PRA2_OBSERVA As String
    OFEX_PRA2_ADVERTENCIA As String
    ' Pruebas asistencia 3
    PRA3_ID As Long
    OFEX_PRA3_PRE_COSTE As Double
    OFEX_PRA3_IMPORTE As Double
    OFEX_PRA3_DESCRIPCION As String
    OFEX_PRA3_OBSERVA As String
    OFEX_PRA3_ADVERTENCIA As String
    ' Pruebas asistencia 4
    PRA4_ID As Long
    OFEX_PRA4_PRE_COSTE As Double
    OFEX_PRA4_IMPORTE As Double
    OFEX_PRA4_DESCRIPCION As String
    OFEX_PRA4_OBSERVA As String
    OFEX_PRA4_ADVERTENCIA As String
    ' Flags
    ImporteCalculado As Boolean
End Type

'=========================================================
' TIPOS PARA CAT�LOGOS IMPLEMENTADOS (15 originales)
'=========================================================

Public Type tCatalogChasis
    CHAS_ID As Long
    CHAS_MODELO As String
    CHAS_DESCRIPCION As String
    CHAS_ADVERTENCIA As String
    CHAS_PRE_COSTE As Double
End Type

Public Type tCatalogCorreas
    CORR_ID As Long
    CORR_MODELO As String
    CORR_DESCRIPCION As String
    CORR_ADVERTENCIA As String
    CORR_PRE_COSTE As Double
    CORR_ATEX_COSTE As Double
End Type

Public Type tCatalogCabezales
    CABE_ID As Long
    CABE_MODELO As String
    CABE_DESCRIPCION As String
    CABE_ADVERTENCIA As String
    CABE_PRE_COSTE As Double
    CABE_BLOQUE_SAS As String
End Type

Public Type tCatalogCilindros
    CILI_ID As Long
    CILI_MODELO As String
    CILI_DESCRIPCION As String
    CILI_ADVERTENCIA As String
    CILI_PRE_COSTE As Double
    CILI_T2_COSTE As Double
    CILI_CO2_COSTE As Double
    CILI_CPI_COSTE As Double
    CILI_SAS_COSTE As Double
    CILI_EMP_COSTE As Double
End Type

Public Type tCatalogMotores
    MOTO_ID As Long
    MOTO_MODELO As String
    MOTO_DESCRIPCION As String
    MOTO_ADVERTENCIA As String
    MOTO_PRE_COSTE As Double
    MTNO_ID As Long
    MTCL_ID As Long
    MTFB_ID As Long
    MTTN_ID As Long
End Type

Public Type tCatalogGases
    GASE_ID As Long
    GASE_DENOMINACION As String
End Type

Public Type tCatalogRefrigeradores
    REFR_ID As Long
    REFR_MODELO As String
    REFR_DESCRIPCION As String
    REFR_ADVERTENCIA As String
End Type

Public Type tCatalogCalderines
    CALD_ID As Long
    CALD_MODELO As String
    CALD_DESCRIPCION As String
    CALD_ADVERTENCIA As String
End Type

Public Type tCatalogVolumenes
    VOLU_ID As Long
    VOLU_CANTIDAD As String
End Type

Public Type tCatalogAcoplamientosDirecto
    ACOD_ID As Long
    ACOD_MODELO As String
    ACOD_DESCRIPCION As String
    ACOD_ADVERTENCIA As String
    ACOD_PRE_COSTE As Double
End Type

Public Type tCatalogReductores
    REDU_ID As Long
    REDU_MODELO As String
    REDU_DESCRIPCION As String
    REDU_ADVERTENCIA As String
    REDU_PRE_COSTE As Double
    REDU_ATEX_COSTE As Double
End Type

Public Type tCatalogValvulasRetencion
    VARE_ID As Long
    VARE_MODELO As String
    VARE_DESCRIPCION As String
    VARE_ADVERTENCIA As String
    VARE_PRE_COSTE As Double
End Type

Public Type tCatalogElectrovalvulas
    ELEC_ID As Long
    ELEC_MODELO As String
    ELEC_DESCRIPCION As String
    ELEC_ADVERTENCIA As String
    ELEC_PRE_COSTE As Double
End Type

Public Type tCatalogFiltros
    FILT_ID As Long
    FILT_MODELO As String
    FILT_DESCRIPCION As String
    FILT_ADVERTENCIA As String
    FILT_PRE_COSTE As Double
End Type

Public Type tCatalogManometros
    MANO_ID As Long
    MANO_MODELO As String
    MANO_DESCRIPCION As String
    MANO_ADVERTENCIA As String
    MANO_PRE_COSTE As Double
End Type

'=========================================================
' TIPOS PARA CAT�LOGOS NUEVOS - FASE 1 (32 cat�logos)
'=========================================================

'--- GRUPO A: EXTRAS (4) ---

Public Type tCatalogEmbalajes
    EMBA_ID As Long
    EMBA_MODELO As String
    EMBA_DESCRIPCION As String
    EMBA_ADVERTENCIA As String
    EMBA_PRE_COSTE As Double
End Type

Public Type tCatalogTransportes
    TRAN_ID As Long
    TRAN_MODELO As String
    TRAN_DESCRIPCION As String
    TRAN_ADVERTENCIA As String
    TRAN_PRE_COSTE As Double
End Type

Public Type tCatalogPuestasMarcha
    PUMA_ID As Long
    PUMA_MODELO As String
    PUMA_DESCRIPCION As String
    PUMA_ADVERTENCIA As String
    PUMA_PRE_COSTE As Double
End Type

Public Type tCatalogPruebasAsistencia
    PRAS_ID As Long
    PRAS_MODELO As String
    PRAS_DESCRIPCION As String
    PRAS_ADVERTENCIA As String
    PRAS_PRE_COSTE As Double
End Type

'--- GRUPO B: ACCESORIOS B�SICOS (5) ---

Public Type tCatalogAeros
    AERO_ID As Long
    AERO_MODELO As String
    AERO_DESCRIPCION As String
    AERO_ADVERTENCIA As String
    AERO_PRE_COSTE As Double
End Type

Public Type tCatalogCajaLocal
    CALO_ID As Long
    CALO_MODELO As String
    CALO_DESCRIPCION As String
    CALO_ADVERTENCIA As String
    CALO_PRE_COSTE As Double
End Type

Public Type tCatalogGruposEngrase
    GREN_ID As Long
    GREN_MODELO As String
    GREN_DESCRIPCION As String
    GREN_ADVERTENCIA As String
    GREN_PRE_COSTE As Double
End Type

Public Type tCatalogLlavesEntrada
    LLEN_ID As Long
    LLEN_MODELO As String
    LLEN_DESCRIPCION As String
    LLEN_ADVERTENCIA As String
    LLEN_PRE_COSTE As Double
End Type

Public Type tCatalogLlavesSalida
    LLSA_ID As Long
    LLSA_MODELO As String
    LLSA_DESCRIPCION As String
    LLSA_ADVERTENCIA As String
    LLSA_PRE_COSTE As Double
End Type

'--- GRUPO E: CABEZAL (1) ---

Public Type tCatalogNormativasCompresor
    NORC_ID As Long
    NORC_MODELO As String
    NORC_DESCRIPCION As String
    NORC_ADVERTENCIA As String
    NORC_PRE_COSTE As Double
End Type

'--- GRUPO H: OPCIONES (5) ---

Public Type tCatalogValvulasPresion
    VAPR_ID As Long
    VAPR_MODELO As String
    VAPR_DESCRIPCION As String
    VAPR_ADVERTENCIA As String
    VAPR_PRE_COSTE As Double
End Type

Public Type tCatalogEngraseCilindros
    ENCI_ID As Long
    ENCI_MODELO As String
    ENCI_DESCRIPCION As String
    ENCI_ADVERTENCIA As String
    ENCI_PRE_COSTE As Double
End Type

Public Type tCatalogPurgadores
    PURG_ID As Long
    PURG_MODELO As String
    PURG_DESCRIPCION As String
    PURG_ADVERTENCIA As String
    PURG_PRE_COSTE As Double
End Type

Public Type tCatalogResistenciasCalefaccion
    RESC_ID As Long
    RESC_MODELO As String
    RESC_DESCRIPCION As String
    RESC_ADVERTENCIA As String
    RESC_PRE_COSTE As Double
End Type

Public Type tCatalogValvulasReguladoras
    VARG_ID As Long
    VARG_MODELO As String
    VARG_DESCRIPCION As String
    VARG_ADVERTENCIA As String
    VARG_PRE_COSTE As Double
End Type

'--- GRUPO J: DATOS GENERALES (3) ---

Public Type tCatalogPaises
    pais_id As Double
    pais_nombre As String
    pais_iso3166_1n As String
    pais_iso3166_1an2 As String
    pais_iso3166_1an3 As String
End Type

Public Type tCatalogComerciales
    COME_ID As Long
    COME_NOMBRE As String
End Type

Public Type tCatalogOfertantes
    OFTN_ID As Long
    OFTN_NOMBRE As String
End Type

'--- GRUPO C: ARRANCADORES (4 tipos) ---

Public Type tCatalogArrancadoresFuerza
    ARR1_ID As Long
    ARR1_MODELO As String
    ARR1_DESCRIPCION As String
    ARR1_ADVERTENCIA As String
    ARR1_PRE_COSTE As Double
End Type

Public Type tCatalogArrancadoresFuerzaControl
    ARR2_ID As Long
    ARR2_MODELO As String
    ARR2_DESCRIPCION As String
    ARR2_ADVERTENCIA As String
    ARR2_PRE_COSTE As Double
End Type

Public Type tCatalogArrancadoresControl
    ARR3_ID As Long
    ARR3_MODELO As String
    ARR3_DESCRIPCION As String
    ARR3_ADVERTENCIA As String
    ARR3_PRE_COSTE As Double
End Type

Public Type tCatalogArrancadoresFuerzaControlPET
    ARR4_ID As Long
    ARR4_MODELO As String
    ARR4_DESCRIPCION As String
    ARR4_ADVERTENCIA As String
    ARR4_PRE_COSTE As Double
End Type

'--- GRUPO D: INSTRUMENTACI�N (10) ---

Public Type tCatalogTransmisoresTemperatura
    TRAT_ID As Long
    TRAT_MODELO As String
    TRAT_DESCRIPCION As String
    TRAT_ADVERTENCIA As String
    TRAT_PRE_COSTE As Double
End Type

Public Type tCatalogTransmisoresPresion
    TRAP_ID As Long
    TRAP_MODELO As String
    TRAP_DESCRIPCION As String
    TRAP_ADVERTENCIA As String
    TRAP_PRE_COSTE As Double
End Type

Public Type tCatalogTermometros
    TERM_ID As Long
    TERM_MODELO As String
    TERM_DESCRIPCION As String
    TERM_ADVERTENCIA As String
    TERM_PRE_COSTE As Double
End Type

Public Type tCatalogElectrovalvulasRegulacion
    ELER_ID As Long
    ELER_MODELO As String
    ELER_DESCRIPCION As String
    ELER_ADVERTENCIA As String
    ELER_PRE_COSTE As Double
    Campo1 As Long
End Type

Public Type tCatalogSensoresCaidaVastago
    SECV_ID As Long
    SECV_MODELO As String
    SECV_DESCRIPCION As String
    SECV_ADVERTENCIA As String
    SECV_PRE_COSTE As Double
End Type

Public Type tCatalogInterruptoresVibracion
    INTV_ID As Long
    INTV_MODELO As String
    INTV_DESCRIPCION As String
    INTV_ADVERTENCIA As String
    INTV_PRE_COSTE As Double
End Type

Public Type tCatalogInterruptoresNivelAceite
    INTA_ID As Long
    INTA_MODELO As String
    INTA_DESCRIPCION As String
    INTA_ADVERTENCIA As String
    INTA_PRE_COSTE As Double
End Type

Public Type tCatalogNivelesCondensados
    NIVC_ID As Long
    NIVC_MODELO As String
    NIVC_DESCRIPCION As String
    NIVC_ADVERTENCIA As String
    NIVC_PRE_COSTE As Double
End Type

Public Type tCatalogValvulasTermostaticas
    VATE_ID As Long
    VATE_MODELO As String
    VATE_DESCRIPCION As String
    VATE_ADVERTENCIA As String
    VATE_PRE_COSTE As Double
End Type

Public Type tCatalogInstrumentaciones
    INTR_ID As Long
    INTR_MODELO As String
    INTR_DESCRIPCION As String
    INTR_ADVERTENCIA As String
    TRAT_ID As Long
    INTR_TRAT_CANTIDAD As Long
    TRAP_ID As Long
    INTR_TRAP_CANTIDAD As Long
    TERM_ID As Long
    INTR_TERM_CANTIDAD As Long
    MANO_ID As Long
    INTR_MANO_CANTIDAD As Long
    ELER_ID As Long
    INTR_ELER_CANTIDAD As Long
    SECV_ID As Long
    INTR_SECV_CANTIDAD As Long
    INTV_ID As Long
    INTR_INTV_CANTIDAD As Long
    INTA_ID As Long
    INTR_INTA_CANTIDAD As Long
    NIVC_ID As Long
    INTR_NIVC_CANTIDAD As Long
    VATE_ID As Long
    INTR_VATE_CANTIDAD As Long
End Type

