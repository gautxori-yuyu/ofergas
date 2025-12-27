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
Public Sub AgregarOtro(ByVal otro As clsOfertaOtro)
    m_Otros.Add otro
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
        cat = m_Catalogos.ObtenerGases(m_DatosGenerales.GASE_ID)
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

'=========================================================
' MÉTODOS PARA OBTENER NOMBRES DE CATÁLOGOS
'=========================================================
Public Function ObtenerNombrePais() As String
    If m_DatosGenerales.pais_id > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogPaises
        cat = m_Catalogos.ObtenerPaises(m_DatosGenerales.pais_id)
        ObtenerNombrePais = cat.pais_nombre
    Else
        ObtenerNombrePais = ""
    End If
End Function
Public Function ObtenerNombreComercial() As String
    If m_DatosGenerales.COME_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogComerciales
        cat = m_Catalogos.ObtenerComerciales(m_DatosGenerales.COME_ID)
        ObtenerNombreComercial = cat.COME_NOMBRE
    Else
        ObtenerNombreComercial = ""
    End If
End Function
Public Function ObtenerNombreOfertante() As String
    If m_DatosGenerales.OFTN_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogOfertantes
        cat = m_Catalogos.ObtenerOfertantes(m_DatosGenerales.OFTN_ID)
        ObtenerNombreOfertante = cat.OFTN_NOMBRE
    Else
        ObtenerNombreOfertante = ""
    End If
End Function
Public Function ObtenerNombreEmbalaje() As String
    If m_Extras.EMBA_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogEmbalajes
        cat = m_Catalogos.ObtenerEmbalajes(m_Extras.EMBA_ID)
        ObtenerNombreEmbalaje = cat.EMBA_MODELO
    Else
        ObtenerNombreEmbalaje = ""
    End If
End Function
Public Function ObtenerNombreTransporte() As String
    If m_Extras.TRAN_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogTransportes
        cat = m_Catalogos.ObtenerTransportes(m_Extras.TRAN_ID)
        ObtenerNombreTransporte = cat.TRAN_MODELO
    Else
        ObtenerNombreTransporte = ""
    End If
End Function
Public Function ObtenerNombrePuestaMarcha() As String
    If m_Extras.PUMA_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogPuestasMarcha
        cat = m_Catalogos.ObtenerPuestasMarcha(m_Extras.PUMA_ID)
        ObtenerNombrePuestaMarcha = cat.PUMA_MODELO
    Else
        ObtenerNombrePuestaMarcha = ""
    End If
End Function
Public Function ObtenerNombrePruebasAsistencia() As String
    If m_Extras.PRA1_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogPruebasAsistencia
        cat = m_Catalogos.ObtenerPruebasAsistencia(m_Extras.PRA1_ID)
        ObtenerNombrePruebasAsistencia = cat.PRAS_MODELO
    Else
        ObtenerNombrePruebasAsistencia = ""
    End If
End Function
Public Function ObtenerNombreMotor() As String
    If m_Accesorios.MOTO_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogMotores
        cat = m_Catalogos.ObtenerMotores(m_Accesorios.MOTO_ID)
        ObtenerNombreMotor = cat.MOTO_MODELO
    Else
        ObtenerNombreMotor = ""
    End If
End Function
Public Function ObtenerNombreChasis() As String
    If m_Modelo.CHAS_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogChasis
        cat = m_Catalogos.ObtenerChasis(m_Modelo.CHAS_ID)
        ObtenerNombreChasis = cat.CHAS_MODELO
    Else
        ObtenerNombreChasis = ""
    End If
End Function
Public Function ObtenerNombreCorrea() As String
    If m_Modelo.CORR_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogCorreas
        cat = m_Catalogos.ObtenerCorreas(m_Modelo.CORR_ID)
        ObtenerNombreCorrea = cat.CORR_MODELO
    Else
        ObtenerNombreCorrea = ""
    End If
End Function
Public Function ObtenerNombreAcoplamientoDirecto() As String
    If m_Modelo.ACOD_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogAcoplamientosDirecto
        cat = m_Catalogos.ObtenerAcoplamientosDirecto(m_Modelo.ACOD_ID)
        ObtenerNombreAcoplamientoDirecto = cat.ACOD_MODELO
    Else
        ObtenerNombreAcoplamientoDirecto = ""
    End If
End Function
Public Function ObtenerNombreReductor() As String
    If m_Modelo.REDU_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogReductores
        cat = m_Catalogos.ObtenerReductores(m_Modelo.REDU_ID)
        ObtenerNombreReductor = cat.REDU_MODELO
    Else
        ObtenerNombreReductor = ""
    End If
End Function
Public Function ObtenerNombreValvulaRetencion() As String
    If m_Opciones.VARE_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogValvulasRetencion
        cat = m_Catalogos.ObtenerValvulasRetencion(m_Opciones.VARE_ID)
        ObtenerNombreValvulaRetencion = cat.VARE_MODELO
    Else
        ObtenerNombreValvulaRetencion = ""
    End If
End Function
Public Function ObtenerNombreElectrovalvula() As String
    If m_Opciones.ELEC_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogElectrovalvulas
        cat = m_Catalogos.ObtenerElectrovalvulas(m_Opciones.ELEC_ID)
        ObtenerNombreElectrovalvula = cat.ELEC_MODELO
    Else
        ObtenerNombreElectrovalvula = ""
    End If
End Function
Public Function ObtenerNombreFiltro() As String
    If m_Accesorios.FILT_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogFiltros
        cat = m_Catalogos.ObtenerFiltros(m_Accesorios.FILT_ID)
        ObtenerNombreFiltro = cat.FILT_MODELO
    Else
        ObtenerNombreFiltro = ""
    End If
End Function
Public Function ObtenerNombreManometro() As String
    If m_Instrumentacion.MANO_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogManometros
        cat = m_Catalogos.ObtenerManometros(m_Instrumentacion.MANO_ID)
        ObtenerNombreManometro = cat.MANO_MODELO
    Else
        ObtenerNombreManometro = ""
    End If
End Function
Public Function ObtenerNombreRefrigerador1() As String
    If m_Refrigeradores.REF1_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogRefrigeradores
        cat = m_Catalogos.ObtenerRefrigeradores(m_Refrigeradores.REF1_ID)
        ObtenerNombreRefrigerador1 = cat.REFR_MODELO
    Else
        ObtenerNombreRefrigerador1 = ""
    End If
End Function

Public Function ObtenerNombreRefrigerador2() As String
    If m_Refrigeradores.REF2_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogRefrigeradores
        cat = m_Catalogos.ObtenerRefrigeradores(m_Refrigeradores.REF2_ID)
        ObtenerNombreRefrigerador2 = cat.REFR_MODELO
    Else
        ObtenerNombreRefrigerador2 = ""
    End If
End Function

Public Function ObtenerNombreRefrigerador3() As String
    If m_Refrigeradores.REF3_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogRefrigeradores
        cat = m_Catalogos.ObtenerRefrigeradores(m_Refrigeradores.REF3_ID)
        ObtenerNombreRefrigerador3 = cat.REFR_MODELO
    Else
        ObtenerNombreRefrigerador3 = ""
    End If
End Function

Public Function ObtenerNombreRefrigerador4() As String
    If m_Refrigeradores.REF4_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogRefrigeradores
        cat = m_Catalogos.ObtenerRefrigeradores(m_Refrigeradores.REF4_ID)
        ObtenerNombreRefrigerador4 = cat.REFR_MODELO
    Else
        ObtenerNombreRefrigerador4 = ""
    End If
End Function

Public Function ObtenerNombreRefrigerador5() As String
    If m_Refrigeradores.REF5_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogRefrigeradores
        cat = m_Catalogos.ObtenerRefrigeradores(m_Refrigeradores.REF5_ID)
        ObtenerNombreRefrigerador5 = cat.REFR_MODELO
    Else
        ObtenerNombreRefrigerador5 = ""
    End If
End Function

Public Function ObtenerNombreRefrigerador6() As String
    If m_Refrigeradores.REF6_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogRefrigeradores
        cat = m_Catalogos.ObtenerRefrigeradores(m_Refrigeradores.REF6_ID)
        ObtenerNombreRefrigerador6 = cat.REFR_MODELO
    Else
        ObtenerNombreRefrigerador6 = ""
    End If
End Function

' Método con parámetro para seleccionar refrigerador (1-6)
Public Function ObtenerNombreRefrigerador(ByVal numeroEtapa As Long) As String
    If numeroEtapa < 1 Or numeroEtapa > 6 Then
        ObtenerNombreRefrigerador = ""
        Exit Function
    End If
    
    If m_Catalogos Is Nothing Then
        ObtenerNombreRefrigerador = ""
        Exit Function
    End If
    
    Dim refID As Long
    Select Case numeroEtapa
        Case 1: refID = m_Refrigeradores.REF1_ID
        Case 2: refID = m_Refrigeradores.REF2_ID
        Case 3: refID = m_Refrigeradores.REF3_ID
        Case 4: refID = m_Refrigeradores.REF4_ID
        Case 5: refID = m_Refrigeradores.REF5_ID
        Case 6: refID = m_Refrigeradores.REF6_ID
    End Select
    
    If refID > 0 Then
        Dim cat As tCatalogRefrigeradores
        cat = m_Catalogos.ObtenerRefrigeradores(refID)
        ObtenerNombreRefrigerador = cat.REFR_MODELO
    Else
        ObtenerNombreRefrigerador = ""
    End If
End Function


Public Function ObtenerNombreAero() As String
    If m_Accesorios.AERO_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogAeros
        cat = m_Catalogos.ObtenerAeros(m_Accesorios.AERO_ID)
        ObtenerNombreAero = cat.AERO_MODELO
    Else
        ObtenerNombreAero = ""
    End If
End Function
Public Function ObtenerNombreCajaLocal() As String
    If m_Accesorios.CALO_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogCajaLocal
        cat = m_Catalogos.ObtenerCajaLocal(m_Accesorios.CALO_ID)
        ObtenerNombreCajaLocal = cat.CALO_MODELO
    Else
        ObtenerNombreCajaLocal = ""
    End If
End Function
Public Function ObtenerNombreGrupoEngrase() As String
    If m_Accesorios.GREN_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogGruposEngrase
        cat = m_Catalogos.ObtenerGruposEngrase(m_Accesorios.GREN_ID)
        ObtenerNombreGrupoEngrase = cat.GREN_MODELO
    Else
        ObtenerNombreGrupoEngrase = ""
    End If
End Function
Public Function ObtenerNombreLlaveEntrada() As String
    If m_Accesorios.LLEN_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogLlavesEntrada
        cat = m_Catalogos.ObtenerLlavesEntrada(m_Accesorios.LLEN_ID)
        ObtenerNombreLlaveEntrada = cat.LLEN_MODELO
    Else
        ObtenerNombreLlaveEntrada = ""
    End If
End Function
Public Function ObtenerNombreLlaveSalida() As String
    If m_Accesorios.LLSA_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogLlavesSalida
        cat = m_Catalogos.ObtenerLlavesSalida(m_Accesorios.LLSA_ID)
        ObtenerNombreLlaveSalida = cat.LLSA_MODELO
    Else
        ObtenerNombreLlaveSalida = ""
    End If
End Function

Public Function ObtenerNombreArrancador() As String
    ' Un solo arrancador determinado por ARRA_ID y OFAC_ARRA_TIPO
    ' OFAC_ARRA_TIPO indica qué tabla de catálogo usar (1-4)
    
    If m_Accesorios.ARRA_ID <= 0 Or m_Catalogos Is Nothing Then
        ObtenerNombreArrancador = ""
        Exit Function
    End If
    
    Select Case m_Accesorios.OFAC_ARRA_TIPO
        Case 1  ' ArrancadoresFuerza
            Dim cat1 As tCatalogArrancadoresFuerza
            cat1 = m_Catalogos.ObtenerArrancadoresFuerza(m_Accesorios.ARRA_ID)
            ObtenerNombreArrancador = cat1.ARR1_MODELO
            
        Case 2  ' ArrancadoresFuerzaControl
            Dim cat2 As tCatalogArrancadoresFuerzaControl
            cat2 = m_Catalogos.ObtenerArrancadoresFuerzaControl(m_Accesorios.ARRA_ID)
            ObtenerNombreArrancador = cat2.ARR2_MODELO
            
        Case 3  ' ArrancadoresControl
            Dim cat3 As tCatalogArrancadoresControl
            cat3 = m_Catalogos.ObtenerArrancadoresControl(m_Accesorios.ARRA_ID)
            ObtenerNombreArrancador = cat3.ARR3_MODELO
            
        Case 4  ' ArrancadoresFuerzaControlPET
            Dim cat4 As tCatalogArrancadoresFuerzaControlPET
            cat4 = m_Catalogos.ObtenerArrancadoresFuerzaControlPET(m_Accesorios.ARRA_ID)
            ObtenerNombreArrancador = cat4.ARR4_MODELO
            
        Case Else
            ObtenerNombreArrancador = ""
    End Select
End Function
Public Function ObtenerNombreCabezal() As String
    If m_Cabezal.CABE_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogCabezales
        cat = m_Catalogos.ObtenerCabezales(m_Cabezal.CABE_ID)
        ObtenerNombreCabezal = cat.CABE_MODELO
    Else
        ObtenerNombreCabezal = ""
    End If
End Function
Public Function ObtenerNombreNormativaCompresor() As String
    If m_Cabezal.NORC_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogNormativasCompresor
        cat = m_Catalogos.ObtenerNormativasCompresor(m_Cabezal.NORC_ID)
        ObtenerNombreNormativaCompresor = cat.NORC_MODELO
    Else
        ObtenerNombreNormativaCompresor = ""
    End If
End Function
Public Function ObtenerNombreCilindro1() As String
    If m_Cabezal.CIL1_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogCilindros
        cat = m_Catalogos.ObtenerCilindros(m_Cabezal.CIL1_ID)
        ObtenerNombreCilindro1 = cat.CILI_MODELO
    Else
        ObtenerNombreCilindro1 = ""
    End If
End Function
Public Function ObtenerNombreCilindro2() As String
    If m_Cabezal.CIL2_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogCilindros
        cat = m_Catalogos.ObtenerCilindros(m_Cabezal.CIL2_ID)
        ObtenerNombreCilindro2 = cat.CILI_MODELO
    Else
        ObtenerNombreCilindro2 = ""
    End If
End Function
Public Function ObtenerNombreCilindro3() As String
    If m_Cabezal.CIL3_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogCilindros
        cat = m_Catalogos.ObtenerCilindros(m_Cabezal.CIL3_ID)
        ObtenerNombreCilindro3 = cat.CILI_MODELO
    Else
        ObtenerNombreCilindro3 = ""
    End If
End Function
Public Function ObtenerNombreCilindro4() As String
    If m_Cabezal.CIL4_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogCilindros
        cat = m_Catalogos.ObtenerCilindros(m_Cabezal.CIL4_ID)
        ObtenerNombreCilindro4 = cat.CILI_MODELO
    Else
        ObtenerNombreCilindro4 = ""
    End If
End Function
Public Function ObtenerNombreCilindro5() As String
    If m_Cabezal.CIL5_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogCilindros
        cat = m_Catalogos.ObtenerCilindros(m_Cabezal.CIL5_ID)
        ObtenerNombreCilindro5 = cat.CILI_MODELO
    Else
        ObtenerNombreCilindro5 = ""
    End If
End Function
Public Function ObtenerNombreCilindro6() As String
    If m_Cabezal.CIL6_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogCilindros
        cat = m_Catalogos.ObtenerCilindros(m_Cabezal.CIL6_ID)
        ObtenerNombreCilindro6 = cat.CILI_MODELO
    Else
        ObtenerNombreCilindro6 = ""
    End If
End Function
Public Function ObtenerNombreValvulaPresion() As String
    If m_Opciones.VAPR_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogValvulasPresion
        cat = m_Catalogos.ObtenerValvulasPresion(m_Opciones.VAPR_ID)
        ObtenerNombreValvulaPresion = cat.VAPR_MODELO
    Else
        ObtenerNombreValvulaPresion = ""
    End If
End Function
Public Function ObtenerNombreEngraseCilindros() As String
    If m_Opciones.ENCI_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogEngraseCilindros
        cat = m_Catalogos.ObtenerEngraseCilindros(m_Opciones.ENCI_ID)
        ObtenerNombreEngraseCilindros = cat.ENCI_MODELO
    Else
        ObtenerNombreEngraseCilindros = ""
    End If
End Function
Public Function ObtenerNombrePurgador() As String
    If m_Opciones.PURG_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogPurgadores
        cat = m_Catalogos.ObtenerPurgadores(m_Opciones.PURG_ID)
        ObtenerNombrePurgador = cat.PURG_MODELO
    Else
        ObtenerNombrePurgador = ""
    End If
End Function
Public Function ObtenerNombreResistenciaCalefaccion() As String
    If m_Opciones.RESC_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogResistenciasCalefaccion
        cat = m_Catalogos.ObtenerResistenciasCalefaccion(m_Opciones.RESC_ID)
        ObtenerNombreResistenciaCalefaccion = cat.RESC_MODELO
    Else
        ObtenerNombreResistenciaCalefaccion = ""
    End If
End Function
Public Function ObtenerNombreValvulaReguladora() As String
    If m_Opciones.VARG_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogValvulasReguladoras
        cat = m_Catalogos.ObtenerValvulasReguladoras(m_Opciones.VARG_ID)
        ObtenerNombreValvulaReguladora = cat.VARG_MODELO
    Else
        ObtenerNombreValvulaReguladora = ""
    End If
End Function
'Public Function ObtenerNombreCalderin(ByVal index As Long) As String
'    If index > 0 And index <= m_Calderines.Count And Not m_Catalogos Is Nothing Then
'        Dim calderin As tOfertasCalderines
'        calderin = m_Calderines(index)
'        If calderin.CALD_ID > 0 Then
'            Dim cat As tCatalogCalderines
'            cat = m_Catalogos.ObtenerCalderines(calderin.CALD_ID)
'            ObtenerNombreCalderin = cat.CALD_MODELO
'        Else
'            ObtenerNombreCalderin = ""
'        End If
'    Else
'        ObtenerNombreCalderin = ""
'    End If
'End Function

' Depósito de Entrada
'Public Function ObtenerNombreDepositoEntrada() As String
'    If m_Calderines.DEPE_ID > 0 And Not m_Catalogos Is Nothing Then
'        Dim cat As tCatalogDepositosEntrada
'        cat = m_Catalogos.ObtenerDepositosEntrada(m_Calderines.DEPE_ID)
'        ObtenerNombreDepositoEntrada = cat.DEPE_MODELO
'    Else
'        ObtenerNombreDepositoEntrada = ""
'    End If
'End Function
'
'' Depósito Final
'Public Function ObtenerNombreDepositoFinal() As String
'    If m_Calderines.DEPF_ID > 0 And Not m_Catalogos Is Nothing Then
'        Dim cat As tCatalogDepositosFinal
'        cat = m_Catalogos.ObtenerDepositosFinal(m_Calderines.DEPF_ID)
'        ObtenerNombreDepositoFinal = cat.DEPF_MODELO
'    Else
'        ObtenerNombreDepositoFinal = ""
'    End If
'End Function


' DEPÓSITOS - Ya estaban correctos, se mantienen
'Public Function ObtenerModeloDepositoEntrada() As String
'    If m_Calderines.DEPE_ID > 0 And Not m_Catalogos Is Nothing Then
'        Dim cat As tCatalogDepositosEntrada
'        cat = m_Catalogos.ObtenerDepositosEntrada(m_Calderines.DEPE_ID)
'        ObtenerModeloDepositoEntrada = cat.DEPE_MODELO
'    Else
'        ObtenerModeloDepositoEntrada = ""
'    End If
'End Function
'
'Public Function ObtenerModeloDepositoFinal() As String
'    If m_Calderines.DEPF_ID > 0 And Not m_Catalogos Is Nothing Then
'        Dim cat As tCatalogDepositosFinal
'        cat = m_Catalogos.ObtenerDepositosFinal(m_Calderines.DEPF_ID)
'        ObtenerModeloDepositoFinal = cat.DEPF_MODELO
'    Else
'        ObtenerModeloDepositoFinal = ""
'    End If
'End Function


Public Function ObtenerNombreInstrumentacion() As String
    If m_Instrumentacion.INTR_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogInstrumentaciones
        cat = m_Catalogos.ObtenerInstrumentaciones(m_Instrumentacion.INTR_ID)
        ObtenerNombreInstrumentacion = cat.INTR_MODELO
    Else
        ObtenerNombreInstrumentacion = ""
    End If
End Function
Public Function ObtenerNombreTransmisorTemperatura() As String
    If m_Instrumentacion.TRAT_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogTransmisoresTemperatura
        cat = m_Catalogos.ObtenerTransmisoresTemperatura(m_Instrumentacion.TRAT_ID)
        ObtenerNombreTransmisorTemperatura = cat.TRAT_MODELO
    Else
        ObtenerNombreTransmisorTemperatura = ""
    End If
End Function
Public Function ObtenerNombreTransmisorPresion() As String
    If m_Instrumentacion.TRAP_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogTransmisoresPresion
        cat = m_Catalogos.ObtenerTransmisoresPresion(m_Instrumentacion.TRAP_ID)
        ObtenerNombreTransmisorPresion = cat.TRAP_MODELO
    Else
        ObtenerNombreTransmisorPresion = ""
    End If
End Function
Public Function ObtenerNombreTermometro() As String
    If m_Instrumentacion.TERM_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogTermometros
        cat = m_Catalogos.ObtenerTermometros(m_Instrumentacion.TERM_ID)
        ObtenerNombreTermometro = cat.TERM_MODELO
    Else
        ObtenerNombreTermometro = ""
    End If
End Function
Public Function ObtenerNombreElectrovalvulaRegulacion() As String
    If m_Instrumentacion.ELER_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogElectrovalvulasRegulacion
        cat = m_Catalogos.ObtenerElectrovalvulasRegulacion(m_Instrumentacion.ELER_ID)
        ObtenerNombreElectrovalvulaRegulacion = cat.ELER_MODELO
    Else
        ObtenerNombreElectrovalvulaRegulacion = ""
    End If
End Function
Public Function ObtenerNombreSensorCaidaVastago() As String
    If m_Instrumentacion.SECV_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogSensoresCaidaVastago
        cat = m_Catalogos.ObtenerSensoresCaidaVastago(m_Instrumentacion.SECV_ID)
        ObtenerNombreSensorCaidaVastago = cat.SECV_MODELO
    Else
        ObtenerNombreSensorCaidaVastago = ""
    End If
End Function
Public Function ObtenerNombreInterruptorVibracion() As String
    If m_Instrumentacion.INTV_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogInterruptoresVibracion
        cat = m_Catalogos.ObtenerInterruptoresVibracion(m_Instrumentacion.INTV_ID)
        ObtenerNombreInterruptorVibracion = cat.INTV_MODELO
    Else
        ObtenerNombreInterruptorVibracion = ""
    End If
End Function
Public Function ObtenerNombreInterruptorNivelAceite() As String
    If m_Instrumentacion.INTA_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogInterruptoresNivelAceite
        cat = m_Catalogos.ObtenerInterruptoresNivelAceite(m_Instrumentacion.INTA_ID)
        ObtenerNombreInterruptorNivelAceite = cat.INTA_MODELO
    Else
        ObtenerNombreInterruptorNivelAceite = ""
    End If
End Function
Public Function ObtenerNombreNivelCondensados() As String
    If m_Instrumentacion.NIVC_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogNivelesCondensados
        cat = m_Catalogos.ObtenerNivelesCondensados(m_Instrumentacion.NIVC_ID)
        ObtenerNombreNivelCondensados = cat.NIVC_MODELO
    Else
        ObtenerNombreNivelCondensados = ""
    End If
End Function
Public Function ObtenerNombreValvulaTermostatica() As String
    If m_Instrumentacion.VATE_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim cat As tCatalogValvulasTermostaticas
        cat = m_Catalogos.ObtenerValvulasTermostaticas(m_Instrumentacion.VATE_ID)
        ObtenerNombreValvulaTermostatica = cat.VATE_MODELO
    Else
        ObtenerNombreValvulaTermostatica = ""
    End If
End Function
Public Function ObtenerVolumenCalderinAspiracion(ByVal etapa As Long) As String
    If etapa < 1 Or etapa > 6 Or m_Catalogos Is Nothing Then
        ObtenerVolumenCalderinAspiracion = ""
        Exit Function
    End If
    
    Dim volumenID As Long
    Select Case etapa
        Case 1: volumenID = m_Calderines.CVA1_ID
        Case 2: volumenID = m_Calderines.CVA2_ID
        Case 3: volumenID = m_Calderines.CVA3_ID
        Case 4: volumenID = m_Calderines.CVA4_ID
        Case 5: volumenID = m_Calderines.CVA5_ID
        Case 6: volumenID = m_Calderines.CVA6_ID
    End Select
    
    If volumenID > 0 Then
        Dim cat As tCatalogVolumenes
        cat = m_Catalogos.ObtenerVolumenes(volumenID)
        ObtenerVolumenCalderinAspiracion = cat.VOLU_CANTIDAD
    Else
        ObtenerVolumenCalderinAspiracion = ""
    End If
End Function

' VOLÚMENES - Escape
Public Function ObtenerVolumenCalderinEscape(ByVal etapa As Long) As String
    If etapa < 1 Or etapa > 6 Or m_Catalogos Is Nothing Then
        ObtenerVolumenCalderinEscape = ""
        Exit Function
    End If
    
    Dim volumenID As Long
    Select Case etapa
        Case 1: volumenID = m_Calderines.CVE1_ID
        Case 2: volumenID = m_Calderines.CVE2_ID
        Case 3: volumenID = m_Calderines.CVE3_ID
        Case 4: volumenID = m_Calderines.CVE4_ID
        Case 5: volumenID = m_Calderines.CVE5_ID
        Case 6: volumenID = m_Calderines.CVE6_ID
    End Select
    
    If volumenID > 0 Then
        Dim cat As tCatalogVolumenes
        cat = m_Catalogos.ObtenerVolumenes(volumenID)
        ObtenerVolumenCalderinEscape = cat.VOLU_CANTIDAD
    Else
        ObtenerVolumenCalderinEscape = ""
    End If
End Function

' MODELOS - Aspiración
Public Function ObtenerModeloCalderinAspiracion(ByVal etapa As Long) As String
    If etapa < 1 Or etapa > 6 Or m_Catalogos Is Nothing Then
        ObtenerModeloCalderinAspiracion = ""
        Exit Function
    End If
    
    Dim calderinID As Long
    Select Case etapa
        Case 1: calderinID = m_Calderines.CLA1_ID
        Case 2: calderinID = m_Calderines.CLA2_ID
        Case 3: calderinID = m_Calderines.CLA3_ID
        Case 4: calderinID = m_Calderines.CLA4_ID
        Case 5: calderinID = m_Calderines.CLA5_ID
        Case 6: calderinID = m_Calderines.CLA6_ID
    End Select
    
    If calderinID > 0 Then
        Dim cat As tCatalogCalderines
        cat = m_Catalogos.ObtenerCalderines(calderinID)
        ObtenerModeloCalderinAspiracion = cat.CALD_MODELO
    Else
        ObtenerModeloCalderinAspiracion = ""
    End If
End Function

' MODELOS - Escape
Public Function ObtenerModeloCalderinEscape(ByVal etapa As Long) As String
    If etapa < 1 Or etapa > 6 Or m_Catalogos Is Nothing Then
        ObtenerModeloCalderinEscape = ""
        Exit Function
    End If
    
    Dim calderinID As Long
    Select Case etapa
        Case 1: calderinID = m_Calderines.CLE1_ID
        Case 2: calderinID = m_Calderines.CLE2_ID
        Case 3: calderinID = m_Calderines.CLE3_ID
        Case 4: calderinID = m_Calderines.CLE4_ID
        Case 5: calderinID = m_Calderines.CLE5_ID
        Case 6: calderinID = m_Calderines.CLE6_ID
    End Select
    
    If calderinID > 0 Then
        Dim cat As tCatalogCalderines
        cat = m_Catalogos.ObtenerCalderines(calderinID)
        ObtenerModeloCalderinEscape = cat.CALD_MODELO
    Else
        ObtenerModeloCalderinEscape = ""
    End If
End Function

' MATERIALES - Aspiración
Public Function ObtenerMaterialCalderinAspiracion(ByVal etapa As Long) As String
    If etapa < 1 Or etapa > 6 Or m_Catalogos Is Nothing Then
        ObtenerMaterialCalderinAspiracion = ""
        Exit Function
    End If
    
    Dim materialID As Long
    Select Case etapa
        Case 1: materialID = m_Calderines.OFCL_CLA1_MATERIAL
        Case 2: materialID = m_Calderines.OFCL_CLA2_MATERIAL
        Case 3: materialID = m_Calderines.OFCL_CLA3_MATERIAL
        Case 4: materialID = m_Calderines.OFCL_CLA4_MATERIAL
        Case 5: materialID = m_Calderines.OFCL_CLA5_MATERIAL
        Case 6: materialID = m_Calderines.OFCL_CLA6_MATERIAL
    End Select
    
    If materialID > 0 Then
        Dim cat As tCatalogTiposMaterial
        cat = m_Catalogos.ObtenerTiposMaterial(materialID)
        ObtenerMaterialCalderinAspiracion = cat.TIMA_MATERIAL
    Else
        ObtenerMaterialCalderinAspiracion = ""
    End If
End Function

' MATERIALES - Escape
Public Function ObtenerMaterialCalderinEscape(ByVal etapa As Long) As String
    If etapa < 1 Or etapa > 6 Or m_Catalogos Is Nothing Then
        ObtenerMaterialCalderinEscape = ""
        Exit Function
    End If
    
    Dim materialID As Long
    Select Case etapa
        Case 1: materialID = m_Calderines.OFCL_CLE1_MATERIAL
        Case 2: materialID = m_Calderines.OFCL_CLE2_MATERIAL
        Case 3: materialID = m_Calderines.OFCL_CLE3_MATERIAL
        Case 4: materialID = m_Calderines.OFCL_CLE4_MATERIAL
        Case 5: materialID = m_Calderines.OFCL_CLE5_MATERIAL
        Case 6: materialID = m_Calderines.OFCL_CLE6_MATERIAL
    End Select
    
    If materialID > 0 Then
        Dim cat As tCatalogTiposMaterial
        cat = m_Catalogos.ObtenerTiposMaterial(materialID)
        ObtenerMaterialCalderinEscape = cat.TIMA_MATERIAL
    Else
        ObtenerMaterialCalderinEscape = ""
    End If
End Function

' =========================================================================
' MÉTODOS PARA PLANTILLAS/PRECONFIGURACIONES
' =========================================================================

Public Function ObtenerPlantillaCalderines() As String
    If m_Calderines.CALD_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim plantilla As tCatalogCalderinesPredeterminados
        plantilla = m_Catalogos.ObtenerCalderinesPredeterminados(m_Calderines.CALD_ID)
        ObtenerPlantillaCalderines = plantilla.CALP_MODELO
    Else
        ObtenerPlantillaCalderines = ""
    End If
End Function

Public Function ObtenerPlantillaInstrumentacion() As String
    If m_Instrumentacion.INTR_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim plantilla As tCatalogInstrumentaciones
        plantilla = m_Catalogos.ObtenerInstrumentaciones(m_Instrumentacion.INTR_ID)
        ObtenerPlantillaInstrumentacion = plantilla.INTR_MODELO
    Else
        ObtenerPlantillaInstrumentacion = ""
    End If
End Function

Public Function ObtenerPlantillaManoObraCabezal() As String
    If m_ManoObra.MOCA_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim plantilla As tCatalogManoObraCabezales
        plantilla = m_Catalogos.ObtenerManoObraCabezales(m_ManoObra.MOCA_ID)
        ObtenerPlantillaManoObraCabezal = plantilla.MAOC_MODELO
    Else
        ObtenerPlantillaManoObraCabezal = ""
    End If
End Function

Public Function ObtenerPlantillaManoObraIngenieria() As String
    If m_ManoObra.MOIN_ID > 0 And Not m_Catalogos Is Nothing Then
        Dim plantilla As tCatalogManoObraIngenierias
        plantilla = m_Catalogos.ObtenerManoObraIngenierias(m_ManoObra.MOIN_ID)
        ObtenerPlantillaManoObraIngenieria = plantilla.MAOI_MODELO
    Else
        ObtenerPlantillaManoObraIngenieria = ""
    End If
End Function

