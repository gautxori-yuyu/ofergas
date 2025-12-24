VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOferta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================
'@Description: Representa una oferta completa con todas sus tablas hijas
'@Purpose: Encapsular lógica de negocio y cálculos de ofertas
'=========================================================
Option Explicit

'=========================================================
' PROPIEDADES PRIVADAS - DATOS DE LA OFERTA
'=========================================================
Private m_DatosGenerales As tOfertasDatosGenerales
Private m_ManoObra As tOfertasManoObra
Private m_Modelo As tOfertasModelo
Private m_Opciones As tOfertasOpciones
Private m_Cabezal As tOfertasCabezal
Private m_Calderines As tOfertasCalderines
Private m_Refrigeradores As tOfertasRefrigeradores
Private m_Accesorios As tOfertasAccesorios
Private m_Instrumentacion As tOfertasInstrumentacion
Private m_Extras As tOfertasExtras

' Tabla 1:N
Private m_Otros As Collection  ' Of clsOfertaOtro

' Referencia a catálogos (inyectada desde fuera)
Private m_Catalogos As clsCatalogos

' Estados de la oferta
Private m_IsDirty As Boolean
Private m_IsNew As Boolean

' Indicadores de qué tablas existen
Private m_TieneManoObra As Boolean
Private m_TieneModelo As Boolean
Private m_TieneOpciones As Boolean
Private m_TieneCabezal As Boolean
Private m_TieneCalderines As Boolean
Private m_TieneRefrigeradores As Boolean
Private m_TieneAccesorios As Boolean
Private m_TieneInstrumentacion As Boolean
Private m_TieneExtras As Boolean

'=========================================================
' INICIALIZACIÓN
'=========================================================
Private Sub Class_Initialize()
    Set m_Otros = New Collection
    m_IsDirty = False
    m_IsNew = True
    
    ' Inicializar flags
    m_TieneManoObra = False
    m_TieneModelo = False
    m_TieneOpciones = False
    m_TieneCabezal = False
    m_TieneCalderines = False
    m_TieneRefrigeradores = False
    m_TieneAccesorios = False
    m_TieneInstrumentacion = False
    m_TieneExtras = False
End Sub

'=========================================================
' PROPIEDADES PÚBLICAS - GET/LET
'=========================================================

Public Property Get DatosGenerales() As tOfertasDatosGenerales
    DatosGenerales = m_DatosGenerales
End Property

Public Property Let DatosGenerales(dg As tOfertasDatosGenerales)
    m_DatosGenerales = dg
    m_IsDirty = True
End Property

Public Property Get ManoObra() As tOfertasManoObra
    ManoObra = m_ManoObra
End Property

Public Property Let ManoObra(mo As tOfertasManoObra)
    m_ManoObra = mo
    m_TieneManoObra = True
    m_IsDirty = True
End Property

Public Property Get Modelo() As tOfertasModelo
    Modelo = m_Modelo
End Property

Public Property Let Modelo(md As tOfertasModelo)
    m_Modelo = md
    m_TieneModelo = True
    m_IsDirty = True
End Property

Public Property Get Opciones() As tOfertasOpciones
    Opciones = m_Opciones
End Property

Public Property Let Opciones(op As tOfertasOpciones)
    m_Opciones = op
    m_TieneOpciones = True
    m_IsDirty = True
End Property

Public Property Get Cabezal() As tOfertasCabezal
    Cabezal = m_Cabezal
End Property

Public Property Let Cabezal(cb As tOfertasCabezal)
    m_Cabezal = cb
    m_TieneCabezal = True
    m_IsDirty = True
End Property

Public Property Get Calderines() As tOfertasCalderines
    Calderines = m_Calderines
End Property

Public Property Let Calderines(cl As tOfertasCalderines)
    m_Calderines = cl
    m_TieneCalderines = True
    m_IsDirty = True
End Property

Public Property Get Refrigeradores() As tOfertasRefrigeradores
    Refrigeradores = m_Refrigeradores
End Property

Public Property Let Refrigeradores(rf As tOfertasRefrigeradores)
    m_Refrigeradores = rf
    m_TieneRefrigeradores = True
    m_IsDirty = True
End Property

Public Property Get Accesorios() As tOfertasAccesorios
    Accesorios = m_Accesorios
End Property

Public Property Let Accesorios(ac As tOfertasAccesorios)
    m_Accesorios = ac
    m_TieneAccesorios = True
    m_IsDirty = True
End Property

Public Property Get Instrumentacion() As tOfertasInstrumentacion
    Instrumentacion = m_Instrumentacion
End Property

Public Property Let Instrumentacion(inst As tOfertasInstrumentacion)
    m_Instrumentacion = inst
    m_TieneInstrumentacion = True
    m_IsDirty = True
End Property

Public Property Get Extras() As tOfertasExtras
    Extras = m_Extras
End Property

Public Property Let Extras(ex As tOfertasExtras)
    m_Extras = ex
    m_TieneExtras = True
    m_IsDirty = True
End Property

Public Property Get Otros() As Collection
    Set Otros = m_Otros
End Property

Public Sub SetCatalogos(ByVal cat As clsCatalogos)
    Set m_Catalogos = cat
End Sub

'=========================================================
' MÉTODOS PÚBLICOS - AGREGAR/ELIMINAR
'=========================================================
'TODO: problema al agregar items, AUNQUE LEE BIEN DE LA DB, SIEMPRE AÑADE EL MISMO REGISTRO...
Public Sub AgregarOtro(ByVal Otro As clsOfertaOtro)
    m_Otros.Add Otro
    m_IsDirty = True
End Sub

'=========================================================
' PROPIEDADES CALCULADAS - ACCESO A CATÁLOGOS
'=========================================================

'@Description: Obtiene la descripción del chasis desde el catálogo o el override
Public Function ObtenerDescripcionChasis() As String
    On Error GoTo ErrHandler
    
    If Not m_TieneModelo Then
        ObtenerDescripcionChasis = ""
        Exit Function
    End If
    
    ' Si hay override, usar ese
    If Len(m_Modelo.OFMO_CHAS_DESCRIPCION) > 0 Then
        ObtenerDescripcionChasis = m_Modelo.OFMO_CHAS_DESCRIPCION
    Else
        ' Si no, consultar catálogo
        If m_Modelo.CHAS_ID > 0 And Not m_Catalogos Is Nothing Then
            Dim cat As tCatalogChasis
            cat = m_Catalogos.ObtenerChasis(m_Modelo.CHAS_ID)
            ObtenerDescripcionChasis = cat.CHAS_DESCRIPCION
        Else
            ObtenerDescripcionChasis = ""
        End If
    End If
    Exit Function
    
ErrHandler:
    ObtenerDescripcionChasis = ""
End Function

'@Description: Obtiene la denominación del gas
Public Function ObtenerNombreGas() As String
    On Error GoTo ErrHandler
    
    If m_DatosGenerales.GASE_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogGases
        cat = m_Catalogos.ObtenerGas(m_DatosGenerales.GASE_ID)
        ObtenerNombreGas = cat.GASE_DENOMINACION
    Else
        ObtenerNombreGas = ""
    End If
    Exit Function
    
ErrHandler:
    ObtenerNombreGas = ""
End Function

'@Description: Calcula el importe total de la oferta
'FIXME: NO ESTA BIEN, HAY QUE APLICAR CIERTOS COEFICIENTES / CRITERIOS...: HABRA QUE DEDUCIR LOS CALCULOS, COMPARANDO CON LAS HOJAS Q GENERA OFERGAS.
Public Function CalcularTotal() As Double
    On Error GoTo ErrHandler
    
    Dim total As Double
    total = 0
    
    ' Sumar mano de obra
    If m_TieneManoObra Then
        If m_ManoObra.ImporteCalculado Then
            total = total + m_ManoObra.OFMA_IMPORTE
        Else
            total = total + CalcularImporteManoObra()
        End If
    End If
    
    ' Sumar modelo
    If m_TieneModelo Then
        total = total + m_Modelo.OFMO_CHAS_IMPORTE
        total = total + m_Modelo.OFMO_CORR_IMPORTE
        total = total + m_Modelo.OFMO_ACOD_IMPORTE
        total = total + m_Modelo.OFMO_REDU_IMPORTE
        total = total + m_Modelo.OFMO_TUDI_IMPORTE
        total = total + m_Modelo.OFMO_TUAG_IMPORTE
    End If
    
    ' Sumar opciones
    If m_TieneOpciones Then
        total = total + m_Opciones.OFOP_VARE_IMPORTE
        total = total + m_Opciones.OFOP_ELEC_IMPORTE
        total = total + m_Opciones.OFOP_VAPR_IMPORTE
        total = total + m_Opciones.OFOP_ENCI_IMPORTE
        total = total + m_Opciones.OFOP_PURG_IMPORTE
        total = total + m_Opciones.OFOP_RESC_IMPORTE
        total = total + m_Opciones.OFOP_VARG_IMPORTE
    End If
    
    ' Sumar cabezal
    If m_TieneCabezal Then
        total = total + m_Cabezal.OFCA_CABE_IMPORTE
        total = total + m_Cabezal.OFCA_CIL1_IMPORTE
        total = total + m_Cabezal.OFCA_CIL2_IMPORTE
        total = total + m_Cabezal.OFCA_CIL3_IMPORTE
        total = total + m_Cabezal.OFCA_CIL4_IMPORTE
        total = total + m_Cabezal.OFCA_CIL5_IMPORTE
        total = total + m_Cabezal.OFCA_CIL6_IMPORTE
        total = total + m_Cabezal.OFCA_BSAS_IMPORTE
        total = total + m_Cabezal.OFCA_NORC_IMPORTE
    End If
    
    ' Sumar calderines
    If m_TieneCalderines Then
        total = total + m_Calderines.OFCL_ETP1_IMPORTE
        total = total + m_Calderines.OFCL_ETP2_IMPORTE
        total = total + m_Calderines.OFCL_ETP3_IMPORTE
        total = total + m_Calderines.OFCL_ETP4_IMPORTE
        total = total + m_Calderines.OFCL_ETP5_IMPORTE
        total = total + m_Calderines.OFCL_ETP6_IMPORTE
        total = total + m_Calderines.OFCL_DEPE_IMPORTE
        total = total + m_Calderines.OFCL_DEPF_IMPORTE
    End If
    
    ' Sumar refrigeradores
    If m_TieneRefrigeradores Then
        total = total + m_Refrigeradores.OFRF_REF1_IMPORTE
        total = total + m_Refrigeradores.OFRF_REF2_IMPORTE
        total = total + m_Refrigeradores.OFRF_REF3_IMPORTE
        total = total + m_Refrigeradores.OFRF_REF4_IMPORTE
        total = total + m_Refrigeradores.OFRF_REF5_IMPORTE
        total = total + m_Refrigeradores.OFRF_REF6_IMPORTE
        total = total + m_Refrigeradores.OFRF_VASE_IMPORTE
    End If
    
    ' Sumar accesorios
    If m_TieneAccesorios Then
        total = total + m_Accesorios.OFAC_MOTO_IMPORTE
        total = total + m_Accesorios.OFAC_MOTO_ESP_IMPORTE
        total = total + m_Accesorios.OFAC_ARRA_IMPORTE
        total = total + m_Accesorios.OFAC_CALO_IMPORTE
        total = total + m_Accesorios.OFAC_FILT_IMPORTE
        total = total + m_Accesorios.OFAC_AERO_IMPORTE
        total = total + m_Accesorios.OFAC_LLEN_IMPORTE
        total = total + m_Accesorios.OFAC_LLSA_IMPORTE
        total = total + m_Accesorios.OFAC_GREN_IMPORTE
    End If
    
    ' Sumar instrumentación
    If m_TieneInstrumentacion Then
        total = total + m_Instrumentacion.OFIN_TRAT_IMPORTE
        total = total + m_Instrumentacion.OFIN_TRAP_IMPORTE
        total = total + m_Instrumentacion.OFIN_TERM_IMPORTE
        total = total + m_Instrumentacion.OFIN_MANO_IMPORTE
        total = total + m_Instrumentacion.OFIN_ELER_IMPORTE
        total = total + m_Instrumentacion.OFIN_SECV_IMPORTE
        total = total + m_Instrumentacion.OFIN_INTV_IMPORTE
        total = total + m_Instrumentacion.OFIN_INTA_IMPORTE
        total = total + m_Instrumentacion.OFIN_NIVC_IMPORTE
        total = total + m_Instrumentacion.OFIN_VATE_IMPORTE
        total = total + m_Instrumentacion.OFIN_INS1_IMPORTE
        total = total + m_Instrumentacion.OFIN_INS2_IMPORTE
        total = total + m_Instrumentacion.OFIN_INS3_IMPORTE
        total = total + m_Instrumentacion.OFIN_INS4_IMPORTE
        total = total + m_Instrumentacion.OFIN_INS5_IMPORTE
        total = total + m_Instrumentacion.OFIN_INS6_IMPORTE
        total = total + m_Instrumentacion.OFIN_INS7_IMPORTE
    End If
    
    ' Sumar extras
    If m_TieneExtras Then
        total = total + m_Extras.OFEX_EMBA_IMPORTE
        total = total + m_Extras.OFEX_TRAN_IMPORTE
        total = total + m_Extras.OFEX_PUMA_IMPORTE
        total = total + m_Extras.OFEX_PRA1_IMPORTE
        total = total + m_Extras.OFEX_PRA2_IMPORTE
        total = total + m_Extras.OFEX_PRA3_IMPORTE
        total = total + m_Extras.OFEX_PRA4_IMPORTE
    End If
    
    ' Sumar otros
    Dim ot As clsOfertaOtro
    For Each ot In m_Otros
        total = total + ot.OFOT_PRE_COSTE
    Next ot
    
    CalcularTotal = total
    Exit Function
    
ErrHandler:
    CalcularTotal = 0
End Function

'@Description: Calcula el importe de mano de obra
Private Function CalcularImporteManoObra() As Double
    Dim importe As Double
    importe = 0
    
    With m_ManoObra
        importe = importe + (.OFMA_MOCA_FASE1_HORAS * .OFMA_MOCA_FASE1_PRECIO)
        importe = importe + (.OFMA_MOCA_FASE2_HORAS * .OFMA_MOCA_FASE2_PRECIO)
        importe = importe + (.OFMA_MOCA_SOLDA_HORAS * .OFMA_MOCA_SOLDA_PRECIO)
        importe = importe + (.OFMA_MOCA_PROBA_HORAS * .OFMA_MOCA_PROBA_PRECIO)
        importe = importe + (.OFMA_MOCA_PINTU_HORAS * .OFMA_MOCA_PINTU_PRECIO)
        importe = importe + (.OFMA_MOCA_ELECT_HORAS * .OFMA_MOCA_ELECT_PRECIO)
        importe = importe + (.OFMA_MOIN_INGEN_HORAS * .OFMA_MOIN_INGEN_PRECIO)
    End With
    
    CalcularImporteManoObra = importe
End Function

'=========================================================
' PROPIEDADES DE ESTADO
'=========================================================

Public Property Get IsDirty() As Boolean
    IsDirty = m_IsDirty
End Property

Public Property Get IsNew() As Boolean
    IsNew = m_IsNew
End Property

Public Property Let IsNew(ByVal Value As Boolean)
    m_IsNew = Value
End Property

'@Description: Indica si la oferta tiene todas las tablas hijas requeridas
Public Function EstaCompleta() As Boolean
    EstaCompleta = m_TieneManoObra And m_TieneModelo And m_TieneOpciones And _
                   m_TieneCabezal And m_TieneCalderines And m_TieneRefrigeradores And _
                   m_TieneAccesorios And m_TieneInstrumentacion And m_TieneExtras And _
                   (m_Otros.Count > 0)
End Function

'@Description: Devuelve un array con las tablas que faltan
Public Function TablasFaltantes() As Variant
    Dim tablas() As String
    Dim i As Integer
    i = 0
    
    ReDim tablas(0 To 9)
    
    If Not m_TieneManoObra Then tablas(i) = "ManoObra": i = i + 1
    If Not m_TieneModelo Then tablas(i) = "Modelo": i = i + 1
    If Not m_TieneOpciones Then tablas(i) = "Opciones": i = i + 1
    If Not m_TieneCabezal Then tablas(i) = "Cabezal": i = i + 1
    If Not m_TieneCalderines Then tablas(i) = "Calderines": i = i + 1
    If Not m_TieneRefrigeradores Then tablas(i) = "Refrigeradores": i = i + 1
    If Not m_TieneAccesorios Then tablas(i) = "Accesorios": i = i + 1
    If Not m_TieneInstrumentacion Then tablas(i) = "Instrumentacion": i = i + 1
    If Not m_TieneExtras Then tablas(i) = "Extras": i = i + 1
    If m_Otros.Count = 0 Then tablas(i) = "Otros": i = i + 1
    
    If i > 0 Then
        ReDim Preserve tablas(0 To i - 1)
        TablasFaltantes = tablas
    Else
        TablasFaltantes = Array()
    End If
End Function
