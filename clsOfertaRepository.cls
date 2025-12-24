VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOfertaRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================
'@Description: Repositorio para lectura/escritura de ofertas
'@Purpose: Capa de acceso a datos optimizada con catálogos
'@NOTE: Este archivo contiene versiones ABREVIADAS de algunos métodos
'       Debes completarlas copiando desde clsOfertaRepository_SECCIONES.txt
'=========================================================
Option Explicit

Private m_Ctx As clsDBContext
Private m_Catalogos As clsCatalogos

Public Sub SetDBContext(ByVal ctx As clsDBContext)
    Set m_Ctx = ctx
End Sub

Public Sub SetCatalogos(ByVal cat As clsCatalogos)
    Set m_Catalogos = cat
End Sub

'=========================================================
'@Description: Lee todas las ofertas del repositorio
'@Returns: Collection de clsOferta
'=========================================================
Public Function LeerTodas() As Collection
    On Error GoTo ErrHandler

    Dim col As New Collection
    Dim rs As Object
    Dim sql As String
    Dim of As clsOferta
    Dim OferID As String

    sql = "SELECT [OFER_ID] FROM [OfertasDatosGenerales] ORDER BY [OFER_FECHA] DESC, [OFER_NUM_OFERTA]"
    Set rs = m_Ctx.AbrirRecordset(sql)

    Do While Not rs.EOF
        OferID = NormalizarGUID(rs("OFER_ID"))
        Set of = LeerPorOferID(OferID)
        col.Add of
        rs.MoveNext
    Loop

    Set LeerTodas = col
    rs.Close
    Exit Function
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsOfertaRepository.LeerTodas", _
        "Error al leer todas las ofertas: " & Err.Description
End Function

'=========================================================
'@Description: Lee una oferta completa por su ID
'@Param OferID: GUID de la oferta (sin llaves)
'@Returns: clsOferta con todos sus datos
'=========================================================
Public Function LeerPorOferID(ByVal OferID As String) As clsOferta
    On Error GoTo ErrHandler
    
    If OferID = "" Or Not isGUID(OferID) Then
        Err.Raise 5, "clsOfertaRepository.LeerPorOferID", "OferID inválido."
    End If

    Dim of As New clsOferta
    of.SetCatalogos m_Catalogos
    
    ' Cargar tabla maestra
    CargarDatosGenerales of, OferID
    
    ' Cargar 10 tablas hijas (1:1 o 1:0..1)
    CargarManoObra of, OferID
    CargarModelo of, OferID
    CargarOpciones of, OferID
    CargarCabezal of, OferID
    CargarCalderines of, OferID
    CargarRefrigeradores of, OferID
    CargarAccesorios of, OferID
    CargarInstrumentacion of, OferID
    CargarExtras of, OferID
    
    ' Cargar tabla 1:N
    CargarOtros of, OferID
    
    of.IsNew = False
    Set LeerPorOferID = of
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "clsOfertaRepository.LeerPorOferID", _
        "Error: " & Err.Description
End Function

'=========================================================
' MÉTODOS PRIVADOS DE CARGA - UNO POR TABLA
' NOTA: CargarCabezal, CargarCalderines, CargarRefrigeradores,
'       CargarAccesorios, CargarInstrumentacion y CargarExtras
'       están ABREVIADOS. Debes completarlos desde el archivo
'       clsOfertaRepository_SECCIONES.txt
'=========================================================

Private Sub CargarDatosGenerales(ByVal Oferta As clsOferta, ByVal OferID As String)
    On Error GoTo ErrHandler

    Dim rs As Object
    Dim dg As tOfertasDatosGenerales

    Set rs = m_Ctx.AbrirRecordset( _
        "SELECT OFER_ID, OFER_NUM_OFERTA, OFER_FECHA, OFER_CLIENTE, OFER_USUARIO_FINAL, " & _
        "       PAIS_ID, COME_ID, OFER_NUM_CALCULO, GASE_ID, OFER_OBSERVACIONES, " & _
        "       OFER_ESTADO_OFERTA, OFTN_ID, OFER_FEC_ULT_MODIF, " & _
        "       OFER_CALD_EXTRA, OFER_CALD_RADIO, OFER_CALD_ASME, OFER_CALD_INOX, OFER_CALD_SELLO " & _
        "FROM OfertasDatosGenerales WHERE OFER_ID={" & OferID & "}")

    If rs.EOF Then
        rs.Close
        Err.Raise 5, , "Oferta no encontrada"
    End If

    With dg
        .OFER_ID = NormalizarGUID(rs("OFER_ID"))
        .OFER_NUM_OFERTA = NzVBA(rs("OFER_NUM_OFERTA"))
        .OFER_FECHA = NzDate(rs("OFER_FECHA"))
        .OFER_CLIENTE = NzVBA(rs("OFER_CLIENTE"))
        .OFER_USUARIO_FINAL = NzVBA(rs("OFER_USUARIO_FINAL"))
        .pais_id = NzLong(rs("PAIS_ID"))
        .COME_ID = NzLong(rs("COME_ID"))
        .OFER_NUM_CALCULO = NzVBA(rs("OFER_NUM_CALCULO"))
        .GASE_ID = NzLong(rs("GASE_ID"))
        .OFER_OBSERVACIONES = NzVBA(rs("OFER_OBSERVACIONES"))
        .OFER_ESTADO_OFERTA = NzLong(rs("OFER_ESTADO_OFERTA"))
        .OFTN_ID = NzLong(rs("OFTN_ID"))
        .OFER_FEC_ULT_MODIF = NzDate(rs("OFER_FEC_ULT_MODIF"))
        .OFER_CALD_EXTRA = NzDouble(rs("OFER_CALD_EXTRA"))
        .OFER_CALD_RADIO = NzLong(rs("OFER_CALD_RADIO"))
        .OFER_CALD_ASME = NzDouble(rs("OFER_CALD_ASME"))
        .OFER_CALD_INOX = NzDouble(rs("OFER_CALD_INOX"))
        .OFER_CALD_SELLO = NzLong(rs("OFER_CALD_SELLO"))
    End With

    Oferta.DatosGenerales = dg
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "CargarDatosGenerales", "Error: " & Err.Description
End Sub

Private Sub CargarManoObra(ByVal Oferta As clsOferta, ByVal OferID As String)
    On Error GoTo ErrHandler
    
    Dim rs As Object
    Dim mo As tOfertasManoObra
    
    Set rs = m_Ctx.AbrirRecordset( _
        "SELECT * FROM OfertasManoObra WHERE OFER_ID={" & OferID & "}")
    
    If Not rs.EOF Then
        With mo
            .OFMA_ID = NormalizarGUID(rs("OFMA_ID"))
            .OFER_ID = NormalizarGUID(rs("OFER_ID"))
            .MOCA_ID = NzLong(rs("MOCA_ID"))
            .OFMA_MOCA_FASE1_HORAS = NzDouble(rs("OFMA_MOCA_FASE1_HORAS"))
            .OFMA_MOCA_FASE1_PRECIO = NzDouble(rs("OFMA_MOCA_FASE1_PRECIO"))
            .OFMA_MOCA_FASE2_HORAS = NzDouble(rs("OFMA_MOCA_FASE2_HORAS"))
            .OFMA_MOCA_FASE2_PRECIO = NzDouble(rs("OFMA_MOCA_FASE2_PRECIO"))
            .OFMA_MOCA_SOLDA_HORAS = NzDouble(rs("OFMA_MOCA_SOLDA_HORAS"))
            .OFMA_MOCA_SOLDA_PRECIO = NzDouble(rs("OFMA_MOCA_SOLDA_PRECIO"))
            .OFMA_MOCA_PROBA_HORAS = NzDouble(rs("OFMA_MOCA_PROBA_HORAS"))
            .OFMA_MOCA_PROBA_PRECIO = NzDouble(rs("OFMA_MOCA_PROBA_PRECIO"))
            .OFMA_MOCA_PINTU_HORAS = NzDouble(rs("OFMA_MOCA_PINTU_HORAS"))
            .OFMA_MOCA_PINTU_PRECIO = NzDouble(rs("OFMA_MOCA_PINTU_PRECIO"))
            .OFMA_MOCA_ELECT_HORAS = NzDouble(rs("OFMA_MOCA_ELECT_HORAS"))
            .OFMA_MOCA_ELECT_PRECIO = NzDouble(rs("OFMA_MOCA_ELECT_PRECIO"))
            .MOIN_ID = NzLong(rs("MOIN_ID"))
            .OFMA_MOIN_INGEN_HORAS = NzDouble(rs("OFMA_MOIN_INGEN_HORAS"))
            .OFMA_MOIN_INGEN_PRECIO = NzDouble(rs("OFMA_MOIN_INGEN_PRECIO"))
            .OFMA_IMPORTE = NzDouble(rs("OFMA_IMPORTE"))
            .ImporteCalculado = (NzDouble(rs("OFMA_IMPORTE")) > 0)
        End With
        Oferta.ManoObra = mo
    End If
    
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "CargarManoObra", "Error: " & Err.Description
End Sub

Private Sub CargarModelo(ByVal Oferta As clsOferta, ByVal OferID As String)
    On Error GoTo ErrHandler
    
    Dim rs As Object
    Dim tm As tOfertasModelo
    
    Set rs = m_Ctx.AbrirRecordset( _
        "SELECT * FROM OfertasModelo WHERE OFER_ID={" & OferID & "}")
    
    If Not rs.EOF Then
        With tm
            .OFMO_ID = NormalizarGUID(rs("OFMO_ID"))
            .OFER_ID = NormalizarGUID(rs("OFER_ID"))
            .CHAS_ID = NzLong(rs("CHAS_ID"))
            .OFMO_CHAS_PRE_COSTE = NzDouble(rs("OFMO_CHAS_PRE_COSTE"))
            .OFMO_CHAS_IMPORTE = NzDouble(rs("OFMO_CHAS_IMPORTE"))
            .OFMO_CHAS_DESCRIPCION = NzVBA(rs("OFMO_CHAS_DESCRIPCION"))
            .OFMO_CHAS_OBSERVA = NzVBA(rs("OFMO_CHAS_OBSERVA"))
            .OFMO_CHAS_ADVERTENCIA = NzVBA(rs("OFMO_CHAS_ADVERTENCIA"))
            .CORR_ID = NzLong(rs("CORR_ID"))
            .OFMO_CORR_PRE_COSTE = NzDouble(rs("OFMO_CORR_PRE_COSTE"))
            .OFMO_CORR_IMPORTE = NzDouble(rs("OFMO_CORR_IMPORTE"))
            .OFMO_CORR_DESCRIPCION = NzVBA(rs("OFMO_CORR_DESCRIPCION"))
            .OFMO_CORR_OBSERVA = NzVBA(rs("OFMO_CORR_OBSERVA"))
            .OFMO_CORR_ADVERTENCIA = NzVBA(rs("OFMO_CORR_ADVERTENCIA"))
            .ACOD_ID = NzLong(rs("ACOD_ID"))
            .OFMO_ACOD_PRE_COSTE = NzDouble(rs("OFMO_ACOD_PRE_COSTE"))
            .OFMO_ACOD_IMPORTE = NzDouble(rs("OFMO_ACOD_IMPORTE"))
            .OFMO_ACOD_DESCRIPCION = NzVBA(rs("OFMO_ACOD_DESCRIPCION"))
            .OFMO_ACOD_OBSERVA = NzVBA(rs("OFMO_ACOD_OBSERVA"))
            .OFMO_ACOD_ADVERTENCIA = NzVBA(rs("OFMO_ACOD_ADVERTENCIA"))
            .REDU_ID = NzLong(rs("REDU_ID"))
            .OFMO_REDU_PRE_COSTE = NzDouble(rs("OFMO_REDU_PRE_COSTE"))
            .OFMO_REDU_IMPORTE = NzDouble(rs("OFMO_REDU_IMPORTE"))
            .OFMO_REDU_DESCRIPCION = NzVBA(rs("OFMO_REDU_DESCRIPCION"))
            .OFMO_REDU_OBSERVA = NzVBA(rs("OFMO_REDU_OBSERVA"))
            .OFMO_REDU_ADVERTENCIA = NzVBA(rs("OFMO_REDU_ADVERTENCIA"))
            .OFMO_TRAN_TIPO = NzLong(rs("OFMO_TRAN_TIPO"))
            .OFMO_ATEX_SN = NzVBA(rs("OFMO_ATEX_SN"))
            .OFMO_ATEX_PRE_COSTE = NzDouble(rs("OFMO_ATEX_PRE_COSTE"))
            .TUAI_ID = NzLong(rs("TUAI_ID"))
            .TUDI_ID = NzLong(rs("TUDI_ID"))
            .OFMO_TUDI_PRE_COSTE = NzDouble(rs("OFMO_TUDI_PRE_COSTE"))
            .OFMO_TUDI_IMPORTE = NzDouble(rs("OFMO_TUDI_IMPORTE"))
            .OFMO_TUDI_DESCRIPCION = NzVBA(rs("OFMO_TUDI_DESCRIPCION"))
            .OFMO_TUDI_OBSERVA = NzVBA(rs("OFMO_TUDI_OBSERVA"))
            .OFMO_TUDI_ADVERTENCIA = NzVBA(rs("OFMO_TUDI_ADVERTENCIA"))
            .OFMO_TUDI_CANTIDAD = NzLong(rs("OFMO_TUDI_CANTIDAD"))
            .TUAG_ID = NzLong(rs("TUAG_ID"))
            .OFMO_TUAG_PRE_COSTE = NzDouble(rs("OFMO_TUAG_PRE_COSTE"))
            .OFMO_TUAG_IMPORTE = NzDouble(rs("OFMO_TUAG_IMPORTE"))
            .OFMO_TUAG_DESCRIPCION = NzVBA(rs("OFMO_TUAG_DESCRIPCION"))
            .OFMO_TUAG_OBSERVA = NzVBA(rs("OFMO_TUAG_OBSERVA"))
            .OFMO_TUAG_ADVERTENCIA = NzVBA(rs("OFMO_TUAG_ADVERTENCIA"))
            .ImporteCalculado = False
        End With
        Oferta.Modelo = tm
    End If
    
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "CargarModelo", "Error: " & Err.Description
End Sub

Private Sub CargarOpciones(ByVal Oferta As clsOferta, ByVal OferID As String)
    On Error GoTo ErrHandler
    
    Dim rs As Object
    Dim op As tOfertasOpciones
    
    Set rs = m_Ctx.AbrirRecordset( _
        "SELECT * FROM OfertasOpciones WHERE OFER_ID={" & OferID & "}")
    
    If Not rs.EOF Then
        With op
            .OFOP_ID = NormalizarGUID(rs("OFOP_ID"))
            .OFER_ID = NormalizarGUID(rs("OFER_ID"))
            .VARE_ID = NzLong(rs("VARE_ID"))
            .OFOP_VARE_PRE_COSTE = NzDouble(rs("OFOP_VARE_PRE_COSTE"))
            .OFOP_VARE_IMPORTE = NzDouble(rs("OFOP_VARE_IMPORTE"))
            .OFOP_VARE_DESCRIPCION = NzVBA(rs("OFOP_VARE_DESCRIPCION"))
            .OFOP_VARE_OBSERVA = NzVBA(rs("OFOP_VARE_OBSERVA"))
            .OFOP_VARE_ADVERTENCIA = NzVBA(rs("OFOP_VARE_ADVERTENCIA"))
            .ELEC_ID = NzLong(rs("ELEC_ID"))
            .OFOP_ELEC_PRE_COSTE = NzDouble(rs("OFOP_ELEC_PRE_COSTE"))
            .OFOP_ELEC_IMPORTE = NzDouble(rs("OFOP_ELEC_IMPORTE"))
            .OFOP_ELEC_DESCRIPCION = NzVBA(rs("OFOP_ELEC_DESCRIPCION"))
            .OFOP_ELEC_OBSERVA = NzVBA(rs("OFOP_ELEC_OBSERVA"))
            .OFOP_ELEC_ADVERTENCIA = NzVBA(rs("OFOP_ELEC_ADVERTENCIA"))
            .VAPR_ID = NzLong(rs("VAPR_ID"))
            .OFOP_VAPR_PRE_COSTE = NzDouble(rs("OFOP_VAPR_PRE_COSTE"))
            .OFOP_VAPR_IMPORTE = NzDouble(rs("OFOP_VAPR_IMPORTE"))
            .OFOP_VAPR_DESCRIPCION = NzVBA(rs("OFOP_VAPR_DESCRIPCION"))
            .OFOP_VAPR_OBSERVA = NzVBA(rs("OFOP_VAPR_OBSERVA"))
            .OFOP_VAPR_ADVERTENCIA = NzVBA(rs("OFOP_VAPR_ADVERTENCIA"))
            .ENCI_ID = NzLong(rs("ENCI_ID"))
            .OFOP_ENCI_PRE_COSTE = NzDouble(rs("OFOP_ENCI_PRE_COSTE"))
            .OFOP_ENCI_IMPORTE = NzDouble(rs("OFOP_ENCI_IMPORTE"))
            .OFOP_ENCI_DESCRIPCION = NzVBA(rs("OFOP_ENCI_DESCRIPCION"))
            .OFOP_ENCI_OBSERVA = NzVBA(rs("OFOP_ENCI_OBSERVA"))
            .OFOP_ENCI_ADVERTENCIA = NzVBA(rs("OFOP_ENCI_ADVERTENCIA"))
            .PURG_ID = NzLong(rs("PURG_ID"))
            .OFOP_PURG_PRE_COSTE = NzDouble(rs("OFOP_PURG_PRE_COSTE"))
            .OFOP_PURG_IMPORTE = NzDouble(rs("OFOP_PURG_IMPORTE"))
            .OFOP_PURG_DESCRIPCION = NzVBA(rs("OFOP_PURG_DESCRIPCION"))
            .OFOP_PURG_OBSERVA = NzVBA(rs("OFOP_PURG_OBSERVA"))
            .OFOP_PURG_ADVERTENCIA = NzVBA(rs("OFOP_PURG_ADVERTENCIA"))
            .OFOP_PURG_CANTIDAD = NzInteger(rs("OFOP_PURG_CANTIDAD"))
            .RESC_ID = NzLong(rs("RESC_ID"))
            .OFOP_RESC_PRE_COSTE = NzDouble(rs("OFOP_RESC_PRE_COSTE"))
            .OFOP_RESC_IMPORTE = NzDouble(rs("OFOP_RESC_IMPORTE"))
            .OFOP_RESC_DESCRIPCION = NzVBA(rs("OFOP_RESC_DESCRIPCION"))
            .OFOP_RESC_OBSERVA = NzVBA(rs("OFOP_RESC_OBSERVA"))
            .OFOP_RESC_ADVERTENCIA = NzVBA(rs("OFOP_RESC_ADVERTENCIA"))
            .VARG_ID = NzLong(rs("VARG_ID"))
            .OFOP_VARG_PRE_COSTE = NzDouble(rs("OFOP_VARG_PRE_COSTE"))
            .OFOP_VARG_IMPORTE = NzDouble(rs("OFOP_VARG_IMPORTE"))
            .OFOP_VARG_DESCRIPCION = NzVBA(rs("OFOP_VARG_DESCRIPCION"))
            .OFOP_VARG_OBSERVA = NzVBA(rs("OFOP_VARG_OBSERVA"))
            .OFOP_VARG_ADVERTENCIA = NzVBA(rs("OFOP_VARG_ADVERTENCIA"))
            .ImporteCalculado = False
        End With
        Oferta.Opciones = op
    End If
    
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "CargarOpciones", "Error: " & Err.Description
End Sub

Private Sub CargarCabezal(ByVal Oferta As clsOferta, ByVal OferID As String)
    On Error GoTo ErrHandler
    
    Dim rs As Object
    Dim cb As tOfertasCabezal
    
    Set rs = m_Ctx.AbrirRecordset( _
        "SELECT * FROM OfertasCabezal WHERE OFER_ID={" & OferID & "}")
    
    If Not rs.EOF Then
        With cb
            .OFCA_ID = NormalizarGUID(rs("OFCA_ID"))
            .OFER_ID = NormalizarGUID(rs("OFER_ID"))
            .CABE_ID = NzLong(rs("CABE_ID"))
            .OFCA_CABE_PRE_COSTE = NzDouble(rs("OFCA_CABE_PRE_COSTE"))
            .OFCA_CABE_IMPORTE = NzDouble(rs("OFCA_CABE_IMPORTE"))
            .OFCA_CABE_DESCRIPCION = NzVBA(rs("OFCA_CABE_DESCRIPCION"))
            .OFCA_CABE_OBSERVA = NzVBA(rs("OFCA_CABE_OBSERVA"))
            .OFCA_CABE_ADVERTENCIA = NzVBA(rs("OFCA_CABE_ADVERTENCIA"))
            .OFCA_NUM_ETAPAS = NzLong(rs("OFCA_NUM_ETAPAS"))
            
            ' Cilindro 1
            .CIL1_ID = NzLong(rs("CIL1_ID"))
            .OFCA_CIL1_PRE_COSTE = NzDouble(rs("OFCA_CIL1_PRE_COSTE"))
            .OFCA_CIL1_IMPORTE = NzDouble(rs("OFCA_CIL1_IMPORTE"))
            .OFCA_CIL1_DESCRIPCION = NzVBA(rs("OFCA_CIL1_DESCRIPCION"))
            .OFCA_CIL1_OBSERVA = NzVBA(rs("OFCA_CIL1_OBSERVA"))
            .OFCA_CIL1_ADVERTENCIA = NzVBA(rs("OFCA_CIL1_ADVERTENCIA"))
            .OFCA_CIL1_CANTIDAD = NzLong(rs("OFCA_CIL1_CANTIDAD"))
            .OFCA_CIL1_EMPA_SN = NzVBA(rs("OFCA_CIL1_EMPA_SN"))
            .OFCA_CIL1_EMPA_PRE_COSTE = NzDouble(rs("OFCA_CIL1_EMPA_PRE_COSTE"))
            .OFCA_CIL1_ANSE_PRE_COSTE = NzDouble(rs("OFCA_CIL1_ANSE_PRE_COSTE"))
            .OFCA_CIL1_BSAS_PRE_COSTE = NzDouble(rs("OFCA_CIL1_BSAS_PRE_COSTE"))
            
            ' Cilindro 2
            .CIL2_ID = NzLong(rs("CIL2_ID"))
            .OFCA_CIL2_PRE_COSTE = NzDouble(rs("OFCA_CIL2_PRE_COSTE"))
            .OFCA_CIL2_IMPORTE = NzDouble(rs("OFCA_CIL2_IMPORTE"))
            .OFCA_CIL2_DESCRIPCION = NzVBA(rs("OFCA_CIL2_DESCRIPCION"))
            .OFCA_CIL2_OBSERVA = NzVBA(rs("OFCA_CIL2_OBSERVA"))
            .OFCA_CIL2_ADVERTENCIA = NzVBA(rs("OFCA_CIL2_ADVERTENCIA"))
            .OFCA_CIL2_CANTIDAD = NzLong(rs("OFCA_CIL2_CANTIDAD"))
            .OFCA_CIL2_EMPA_SN = NzVBA(rs("OFCA_CIL2_EMPA_SN"))
            .OFCA_CIL2_EMPA_PRE_COSTE = NzDouble(rs("OFCA_CIL2_EMPA_PRE_COSTE"))
            .OFCA_CIL2_ANSE_PRE_COSTE = NzDouble(rs("OFCA_CIL2_ANSE_PRE_COSTE"))
            .OFCA_CIL2_BSAS_PRE_COSTE = NzDouble(rs("OFCA_CIL2_BSAS_PRE_COSTE"))
            
            ' Cilindro 3
            .CIL3_ID = NzLong(rs("CIL3_ID"))
            .OFCA_CIL3_PRE_COSTE = NzDouble(rs("OFCA_CIL3_PRE_COSTE"))
            .OFCA_CIL3_IMPORTE = NzDouble(rs("OFCA_CIL3_IMPORTE"))
            .OFCA_CIL3_DESCRIPCION = NzVBA(rs("OFCA_CIL3_DESCRIPCION"))
            .OFCA_CIL3_OBSERVA = NzVBA(rs("OFCA_CIL3_OBSERVA"))
            .OFCA_CIL3_ADVERTENCIA = NzVBA(rs("OFCA_CIL3_ADVERTENCIA"))
            .OFCA_CIL3_CANTIDAD = NzLong(rs("OFCA_CIL3_CANTIDAD"))
            .OFCA_CIL3_EMPA_SN = NzVBA(rs("OFCA_CIL3_EMPA_SN"))
            .OFCA_CIL3_EMPA_PRE_COSTE = NzDouble(rs("OFCA_CIL3_EMPA_PRE_COSTE"))
            .OFCA_CIL3_ANSE_PRE_COSTE = NzDouble(rs("OFCA_CIL3_ANSE_PRE_COSTE"))
            .OFCA_CIL3_BSAS_PRE_COSTE = NzDouble(rs("OFCA_CIL3_BSAS_PRE_COSTE"))
            
            ' Cilindro 4
            .CIL4_ID = NzLong(rs("CIL4_ID"))
            .OFCA_CIL4_PRE_COSTE = NzDouble(rs("OFCA_CIL4_PRE_COSTE"))
            .OFCA_CIL4_IMPORTE = NzDouble(rs("OFCA_CIL4_IMPORTE"))
            .OFCA_CIL4_DESCRIPCION = NzVBA(rs("OFCA_CIL4_DESCRIPCION"))
            .OFCA_CIL4_OBSERVA = NzVBA(rs("OFCA_CIL4_OBSERVA"))
            .OFCA_CIL4_ADVERTENCIA = NzVBA(rs("OFCA_CIL4_ADVERTENCIA"))
            .OFCA_CIL4_CANTIDAD = NzLong(rs("OFCA_CIL4_CANTIDAD"))
            .OFCA_CIL4_EMPA_SN = NzVBA(rs("OFCA_CIL4_EMPA_SN"))
            .OFCA_CIL4_EMPA_PRE_COSTE = NzDouble(rs("OFCA_CIL4_EMPA_PRE_COSTE"))
            .OFCA_CIL4_ANSE_PRE_COSTE = NzDouble(rs("OFCA_CIL4_ANSE_PRE_COSTE"))
            .OFCA_CIL4_BSAS_PRE_COSTE = NzDouble(rs("OFCA_CIL4_BSAS_PRE_COSTE"))
            
            ' Cilindro 5
            .CIL5_ID = NzLong(rs("CIL5_ID"))
            .OFCA_CIL5_PRE_COSTE = NzDouble(rs("OFCA_CIL5_PRE_COSTE"))
            .OFCA_CIL5_IMPORTE = NzDouble(rs("OFCA_CIL5_IMPORTE"))
            .OFCA_CIL5_DESCRIPCION = NzVBA(rs("OFCA_CIL5_DESCRIPCION"))
            .OFCA_CIL5_OBSERVA = NzVBA(rs("OFCA_CIL5_OBSERVA"))
            .OFCA_CIL5_ADVERTENCIA = NzVBA(rs("OFCA_CIL5_ADVERTENCIA"))
            .OFCA_CIL5_CANTIDAD = NzLong(rs("OFCA_CIL5_CANTIDAD"))
            .OFCA_CIL5_EMPA_SN = NzVBA(rs("OFCA_CIL5_EMPA_SN"))
            .OFCA_CIL5_EMPA_PRE_COSTE = NzDouble(rs("OFCA_CIL5_EMPA_PRE_COSTE"))
            .OFCA_CIL5_ANSE_PRE_COSTE = NzDouble(rs("OFCA_CIL5_ANSE_PRE_COSTE"))
            .OFCA_CIL5_BSAS_PRE_COSTE = NzDouble(rs("OFCA_CIL5_BSAS_PRE_COSTE"))
            
            ' Cilindro 6
            .CIL6_ID = NzLong(rs("CIL6_ID"))
            .OFCA_CIL6_PRE_COSTE = NzDouble(rs("OFCA_CIL6_PRE_COSTE"))
            .OFCA_CIL6_IMPORTE = NzDouble(rs("OFCA_CIL6_IMPORTE"))
            .OFCA_CIL6_DESCRIPCION = NzVBA(rs("OFCA_CIL6_DESCRIPCION"))
            .OFCA_CIL6_OBSERVA = NzVBA(rs("OFCA_CIL6_OBSERVA"))
            .OFCA_CIL6_ADVERTENCIA = NzVBA(rs("OFCA_CIL6_ADVERTENCIA"))
            .OFCA_CIL6_CANTIDAD = NzLong(rs("OFCA_CIL6_CANTIDAD"))
            .OFCA_CIL6_EMPA_SN = NzVBA(rs("OFCA_CIL6_EMPA_SN"))
            .OFCA_CIL6_EMPA_PRE_COSTE = NzDouble(rs("OFCA_CIL6_EMPA_PRE_COSTE"))
            .OFCA_CIL6_ANSE_PRE_COSTE = NzDouble(rs("OFCA_CIL6_ANSE_PRE_COSTE"))
            .OFCA_CIL6_BSAS_PRE_COSTE = NzDouble(rs("OFCA_CIL6_BSAS_PRE_COSTE"))
            
            ' Anti-surge y bloque SAS
            .OFCA_ANSE_TIPO = NzLong(rs("OFCA_ANSE_TIPO"))
            .OFCA_BSAS_SN = NzVBA(rs("OFCA_BSAS_SN"))
            .OFCA_BSAS_PRE_COSTE = NzDouble(rs("OFCA_BSAS_PRE_COSTE"))
            .OFCA_BSAS_IMPORTE = NzDouble(rs("OFCA_BSAS_IMPORTE"))
            
            ' Normativa compresor
            .NORC_ID = NzLong(rs("NORC_ID"))
            .OFCA_NORC_PRE_COSTE = NzDouble(rs("OFCA_NORC_PRE_COSTE"))
            .OFCA_NORC_IMPORTE = NzDouble(rs("OFCA_NORC_IMPORTE"))
            .OFCA_NORC_DESCRIPCION = NzVBA(rs("OFCA_NORC_DESCRIPCION"))
            .OFCA_NORC_OBSERVA = NzVBA(rs("OFCA_NORC_OBSERVA"))
            .OFCA_NORC_ADVERTENCIA = NzVBA(rs("OFCA_NORC_ADVERTENCIA"))
            
            .ImporteCalculado = False
        End With
        Oferta.Cabezal = cb
    End If
    
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "CargarCabezal", "Error: " & Err.Description
End Sub


Private Sub CargarCalderines(ByVal Oferta As clsOferta, ByVal OferID As String)
                                                                           
    On Error GoTo ErrHandler
    
    Dim rs As Object
    Dim cl As tOfertasCalderines
    
    Set rs = m_Ctx.AbrirRecordset( _
        "SELECT * FROM OfertasCalderines WHERE OFER_ID={" & OferID & "}")
    
    If Not rs.EOF Then
        With cl
            .OFCL_ID = NzLong(rs("OFCL_ID"))
            .OFER_ID = NormalizarGUID(rs("OFER_ID"))
            .CALD_ID = NzLong(rs("CALD_ID"))
            
            ' ========== ETAPA 1 ==========
            ' Volumen aire etapa 1
            .CVA1_ID = NzLong(rs("CVA1_ID"))
            .CLA1_ID = NzLong(rs("CLA1_ID"))
            .OFCL_CLA1_MATERIAL = NzVBA(rs("OFCL_CLA1_MATERIAL"))
            .OFCL_CLA1_PRE_COSTE = NzDouble(rs("OFCL_CLA1_PRE_COSTE"))
            .OFCL_CLA1_DESCRIPCION = NzVBA(rs("OFCL_CLA1_DESCRIPCION"))
            .OFCL_CLA1_OBSERVA = NzVBA(rs("OFCL_CLA1_OBSERVA"))
            .OFCL_CLA1_ADVERTENCIA = NzVBA(rs("OFCL_CLA1_ADVERTENCIA"))
            .OFCL_CLA1_CANTIDAD = NzLong(rs("OFCL_CLA1_CANTIDAD"))
            
            ' Volumen entrada etapa 1
            .CVE1_ID = NzLong(rs("CVE1_ID"))
            .CLE1_ID = NzLong(rs("CLE1_ID"))
            .OFCL_CLE1_MATERIAL = NzVBA(rs("OFCL_CLE1_MATERIAL"))
            .OFCL_CLE1_PRE_COSTE = NzDouble(rs("OFCL_CLE1_PRE_COSTE"))
            .OFCL_CLE1_DESCRIPCION = NzVBA(rs("OFCL_CLE1_DESCRIPCION"))
            .OFCL_CLE1_OBSERVA = NzVBA(rs("OFCL_CLE1_OBSERVA"))
            .OFCL_CLE1_ADVERTENCIA = NzVBA(rs("OFCL_CLE1_ADVERTENCIA"))
            .OFCL_CLE1_CANTIDAD = NzLong(rs("OFCL_CLE1_CANTIDAD"))
            
            .OFCL_ETP1_IMPORTE = NzDouble(rs("OFCL_ETP1_IMPORTE"))
            
            ' ========== ETAPA 2 ==========
            .CVA2_ID = NzLong(rs("CVA2_ID"))
            .CLA2_ID = NzLong(rs("CLA2_ID"))
            .OFCL_CLA2_MATERIAL = NzVBA(rs("OFCL_CLA2_MATERIAL"))
            .OFCL_CLA2_PRE_COSTE = NzDouble(rs("OFCL_CLA2_PRE_COSTE"))
            .OFCL_CLA2_DESCRIPCION = NzVBA(rs("OFCL_CLA2_DESCRIPCION"))
            .OFCL_CLA2_OBSERVA = NzVBA(rs("OFCL_CLA2_OBSERVA"))
            .OFCL_CLA2_ADVERTENCIA = NzVBA(rs("OFCL_CLA2_ADVERTENCIA"))
            .OFCL_CLA2_CANTIDAD = NzLong(rs("OFCL_CLA2_CANTIDAD"))
            
            .CVE2_ID = NzLong(rs("CVE2_ID"))
            .CLE2_ID = NzLong(rs("CLE2_ID"))
            .OFCL_CLE2_MATERIAL = NzVBA(rs("OFCL_CLE2_MATERIAL"))
            .OFCL_CLE2_PRE_COSTE = NzDouble(rs("OFCL_CLE2_PRE_COSTE"))
            .OFCL_CLE2_DESCRIPCION = NzVBA(rs("OFCL_CLE2_DESCRIPCION"))
            .OFCL_CLE2_OBSERVA = NzVBA(rs("OFCL_CLE2_OBSERVA"))
            .OFCL_CLE2_ADVERTENCIA = NzVBA(rs("OFCL_CLE2_ADVERTENCIA"))
            .OFCL_CLE2_CANTIDAD = NzLong(rs("OFCL_CLE2_CANTIDAD"))
            
            .OFCL_ETP2_IMPORTE = NzDouble(rs("OFCL_ETP2_IMPORTE"))
            
            ' ========== ETAPA 3 ==========
            .CVA3_ID = NzLong(rs("CVA3_ID"))
            .CLA3_ID = NzLong(rs("CLA3_ID"))
            .OFCL_CLA3_MATERIAL = NzVBA(rs("OFCL_CLA3_MATERIAL"))
            .OFCL_CLA3_PRE_COSTE = NzDouble(rs("OFCL_CLA3_PRE_COSTE"))
            .OFCL_CLA3_DESCRIPCION = NzVBA(rs("OFCL_CLA3_DESCRIPCION"))
            .OFCL_CLA3_OBSERVA = NzVBA(rs("OFCL_CLA3_OBSERVA"))
            .OFCL_CLA3_ADVERTENCIA = NzVBA(rs("OFCL_CLA3_ADVERTENCIA"))
            .OFCL_CLA3_CANTIDAD = NzLong(rs("OFCL_CLA3_CANTIDAD"))
            
            .CVE3_ID = NzLong(rs("CVE3_ID"))
            .CLE3_ID = NzLong(rs("CLE3_ID"))
            .OFCL_CLE3_MATERIAL = NzVBA(rs("OFCL_CLE3_MATERIAL"))
            .OFCL_CLE3_PRE_COSTE = NzDouble(rs("OFCL_CLE3_PRE_COSTE"))
            .OFCL_CLE3_DESCRIPCION = NzVBA(rs("OFCL_CLE3_DESCRIPCION"))
            .OFCL_CLE3_OBSERVA = NzVBA(rs("OFCL_CLE3_OBSERVA"))
            .OFCL_CLE3_ADVERTENCIA = NzVBA(rs("OFCL_CLE3_ADVERTENCIA"))
            .OFCL_CLE3_CANTIDAD = NzLong(rs("OFCL_CLE3_CANTIDAD"))
            
            .OFCL_ETP3_IMPORTE = NzDouble(rs("OFCL_ETP3_IMPORTE"))
            
            ' ========== ETAPA 4 ==========
            .CVA4_ID = NzLong(rs("CVA4_ID"))
            .CLA4_ID = NzLong(rs("CLA4_ID"))
            .OFCL_CLA4_MATERIAL = NzVBA(rs("OFCL_CLA4_MATERIAL"))
            .OFCL_CLA4_PRE_COSTE = NzDouble(rs("OFCL_CLA4_PRE_COSTE"))
            .OFCL_CLA4_DESCRIPCION = NzVBA(rs("OFCL_CLA4_DESCRIPCION"))
            .OFCL_CLA4_OBSERVA = NzVBA(rs("OFCL_CLA4_OBSERVA"))
            .OFCL_CLA4_ADVERTENCIA = NzVBA(rs("OFCL_CLA4_ADVERTENCIA"))
            .OFCL_CLA4_CANTIDAD = NzLong(rs("OFCL_CLA4_CANTIDAD"))
            
            .CVE4_ID = NzLong(rs("CVE4_ID"))
            .CLE4_ID = NzLong(rs("CLE4_ID"))
            .OFCL_CLE4_MATERIAL = NzVBA(rs("OFCL_CLE4_MATERIAL"))
            .OFCL_CLE4_PRE_COSTE = NzDouble(rs("OFCL_CLE4_PRE_COSTE"))
            .OFCL_CLE4_DESCRIPCION = NzVBA(rs("OFCL_CLE4_DESCRIPCION"))
            .OFCL_CLE4_OBSERVA = NzVBA(rs("OFCL_CLE4_OBSERVA"))
            .OFCL_CLE4_ADVERTENCIA = NzVBA(rs("OFCL_CLE4_ADVERTENCIA"))
            .OFCL_CLE4_CANTIDAD = NzLong(rs("OFCL_CLE4_CANTIDAD"))
            
            .OFCL_ETP4_IMPORTE = NzDouble(rs("OFCL_ETP4_IMPORTE"))
            
            ' ========== ETAPA 5 ==========
            .CVA5_ID = NzLong(rs("CVA5_ID"))
            .CLA5_ID = NzLong(rs("CLA5_ID"))
            .OFCL_CLA5_MATERIAL = NzVBA(rs("OFCL_CLA5_MATERIAL"))
            .OFCL_CLA5_PRE_COSTE = NzDouble(rs("OFCL_CLA5_PRE_COSTE"))
            .OFCL_CLA5_DESCRIPCION = NzVBA(rs("OFCL_CLA5_DESCRIPCION"))
            .OFCL_CLA5_OBSERVA = NzVBA(rs("OFCL_CLA5_OBSERVA"))
            .OFCL_CLA5_ADVERTENCIA = NzVBA(rs("OFCL_CLA5_ADVERTENCIA"))
            .OFCL_CLA5_CANTIDAD = NzLong(rs("OFCL_CLA5_CANTIDAD"))
            
            .CVE5_ID = NzLong(rs("CVE5_ID"))
            .CLE5_ID = NzLong(rs("CLE5_ID"))
            .OFCL_CLE5_MATERIAL = NzVBA(rs("OFCL_CLE5_MATERIAL"))
            .OFCL_CLE5_PRE_COSTE = NzDouble(rs("OFCL_CLE5_PRE_COSTE"))
            .OFCL_CLE5_DESCRIPCION = NzVBA(rs("OFCL_CLE5_DESCRIPCION"))
            .OFCL_CLE5_OBSERVA = NzVBA(rs("OFCL_CLE5_OBSERVA"))
            .OFCL_CLE5_ADVERTENCIA = NzVBA(rs("OFCL_CLE5_ADVERTENCIA"))
            .OFCL_CLE5_CANTIDAD = NzLong(rs("OFCL_CLE5_CANTIDAD"))
            
            .OFCL_ETP5_IMPORTE = NzDouble(rs("OFCL_ETP5_IMPORTE"))
            
            ' ========== ETAPA 6 ==========
            .CVA6_ID = NzLong(rs("CVA6_ID"))
            .CLA6_ID = NzLong(rs("CLA6_ID"))
            .OFCL_CLA6_MATERIAL = NzVBA(rs("OFCL_CLA6_MATERIAL"))
            .OFCL_CLA6_PRE_COSTE = NzDouble(rs("OFCL_CLA6_PRE_COSTE"))
            .OFCL_CLA6_DESCRIPCION = NzVBA(rs("OFCL_CLA6_DESCRIPCION"))
            .OFCL_CLA6_OBSERVA = NzVBA(rs("OFCL_CLA6_OBSERVA"))
            .OFCL_CLA6_ADVERTENCIA = NzVBA(rs("OFCL_CLA6_ADVERTENCIA"))
            .OFCL_CLA6_CANTIDAD = NzLong(rs("OFCL_CLA6_CANTIDAD"))
            
            .CVE6_ID = NzLong(rs("CVE6_ID"))
            .CLE6_ID = NzLong(rs("CLE6_ID"))
            .OFCL_CLE6_MATERIAL = NzVBA(rs("OFCL_CLE6_MATERIAL"))
            .OFCL_CLE6_PRE_COSTE = NzDouble(rs("OFCL_CLE6_PRE_COSTE"))
            .OFCL_CLE6_DESCRIPCION = NzVBA(rs("OFCL_CLE6_DESCRIPCION"))
            .OFCL_CLE6_OBSERVA = NzVBA(rs("OFCL_CLE6_OBSERVA"))
            .OFCL_CLE6_ADVERTENCIA = NzVBA(rs("OFCL_CLE6_ADVERTENCIA"))
            .OFCL_CLE6_CANTIDAD = NzLong(rs("OFCL_CLE6_CANTIDAD"))
            
            .OFCL_ETP6_IMPORTE = NzDouble(rs("OFCL_ETP6_IMPORTE"))
            
            ' ========== OPCIONES ADICIONALES ==========
            .OFCL_EXTRA_SN = NzVBA(rs("OFCL_EXTRA_SN"))
            .OFCL_RADIO_SN = NzVBA(rs("OFCL_RADIO_SN"))
            .OFCL_ASME_SN = NzVBA(rs("OFCL_ASME_SN"))
            .OFCL_SELLOU_SN = NzVBA(rs("OFCL_SELLOU_SN"))
            .OFCL_SELLOU_IMPORTE = NzLong(rs("OFCL_SELLOU_IMPORTE"))
            
            ' ========== DEPÓSITOS ENTRADA ==========
            .DEPE_ID = NzLong(rs("DEPE_ID"))
            .OFCL_DEPE_PRE_COSTE = NzDouble(rs("OFCL_DEPE_PRE_COSTE"))
            .OFCL_DEPE_IMPORTE = NzDouble(rs("OFCL_DEPE_IMPORTE"))
            .OFCL_DEPE_DESCRIPCION = NzVBA(rs("OFCL_DEPE_DESCRIPCION"))
            .OFCL_DEPE_OBSERVA = NzVBA(rs("OFCL_DEPE_OBSERVA"))
            .OFCL_DEPE_ADVERTENCIA = NzVBA(rs("OFCL_DEPE_ADVERTENCIA"))
            
            ' ========== DEPÓSITOS FINAL ==========
            .DEPF_ID = NzLong(rs("DEPF_ID"))
            .OFCL_DEPF_PRE_COSTE = NzDouble(rs("OFCL_DEPF_PRE_COSTE"))
            .OFCL_DEPF_IMPORTE = NzDouble(rs("OFCL_DEPF_IMPORTE"))
            .OFCL_DEPF_DESCRIPCION = NzVBA(rs("OFCL_DEPF_DESCRIPCION"))
            .OFCL_DEPF_OBSERVA = NzVBA(rs("OFCL_DEPF_OBSERVA"))
            .OFCL_DEPF_ADVERTENCIA = NzVBA(rs("OFCL_DEPF_ADVERTENCIA"))
            
            .ImporteCalculado = False
        End With
        Oferta.Calderines = cl
    End If
    
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "CargarCalderines", "Error: " & Err.Description
End Sub

Private Sub CargarRefrigeradores(ByVal Oferta As clsOferta, ByVal OferID As String)
                                                                           
    On Error GoTo ErrHandler
    
    Dim rs As Object
    Dim rf As tOfertasRefrigeradores
    
    Set rs = m_Ctx.AbrirRecordset( _
        "SELECT * FROM OfertasRefrigeradores WHERE OFER_ID={" & OferID & "}")
    
    If Not rs.EOF Then
        With rf
            .OFRF_ID = NormalizarGUID(rs("OFRF_ID"))
            .OFER_ID = NormalizarGUID(rs("OFER_ID"))
            .TITU_ID = NzLong(rs("TITU_ID"))
            
            ' ========== REFRIGERADOR 1 ==========
            .REF1_ID = NzLong(rs("REF1_ID"))
            .OFRF_REF1_TIPO = NzLong(rs("OFRF_REF1_TIPO"))
            .OFRF_REF1_PN = NzLong(rs("OFRF_REF1_PN"))
            .OFRF_REF1_CE = NzLong(rs("OFRF_REF1_CE"))
            .OFRF_REF1_CS = NzLong(rs("OFRF_REF1_CS"))
            .OFRF_REF1_PRE_COSTE = NzDouble(rs("OFRF_REF1_PRE_COSTE"))
            .OFRF_REF1_CE_COSTE = NzDouble(rs("OFRF_REF1_CE_COSTE"))
            .OFRF_REF1_CS_COSTE = NzDouble(rs("OFRF_REF1_CS_COSTE"))
            .OFRF_REF1_IMPORTE = NzDouble(rs("OFRF_REF1_IMPORTE"))
            .OFRF_REF1_DESCRIPCION = NzVBA(rs("OFRF_REF1_DESCRIPCION"))
            .OFRF_REF1_OBSERVA = NzVBA(rs("OFRF_REF1_OBSERVA"))
            .OFRF_REF1_ADVERTENCIA = NzVBA(rs("OFRF_REF1_ADVERTENCIA"))
            .OFRF_REF1_CANTIDAD = NzLong(rs("OFRF_REF1_CANTIDAD"))
            
            ' ========== REFRIGERADOR 2 ==========
            .REF2_ID = NzLong(rs("REF2_ID"))
            .OFRF_REF2_TIPO = NzLong(rs("OFRF_REF2_TIPO"))
            .OFRF_REF2_PN = NzLong(rs("OFRF_REF2_PN"))
            .OFRF_REF2_CE = NzLong(rs("OFRF_REF2_CE"))
            .OFRF_REF2_CS = NzLong(rs("OFRF_REF2_CS"))
            .OFRF_REF2_PRE_COSTE = NzDouble(rs("OFRF_REF2_PRE_COSTE"))
            .OFRF_REF2_CE_COSTE = NzDouble(rs("OFRF_REF2_CE_COSTE"))
            .OFRF_REF2_CS_COSTE = NzDouble(rs("OFRF_REF2_CS_COSTE"))
            .OFRF_REF2_IMPORTE = NzDouble(rs("OFRF_REF2_IMPORTE"))
            .OFRF_REF2_DESCRIPCION = NzVBA(rs("OFRF_REF2_DESCRIPCION"))
            .OFRF_REF2_OBSERVA = NzVBA(rs("OFRF_REF2_OBSERVA"))
            .OFRF_REF2_ADVERTENCIA = NzVBA(rs("OFRF_REF2_ADVERTENCIA"))
            .OFRF_REF2_CANTIDAD = NzLong(rs("OFRF_REF2_CANTIDAD"))
            
            ' ========== REFRIGERADOR 3 ==========
            .REF3_ID = NzLong(rs("REF3_ID"))
            .OFRF_REF3_TIPO = NzLong(rs("OFRF_REF3_TIPO"))
            .OFRF_REF3_PN = NzLong(rs("OFRF_REF3_PN"))
            .OFRF_REF3_CE = NzLong(rs("OFRF_REF3_CE"))
            .OFRF_REF3_CS = NzLong(rs("OFRF_REF3_CS"))
            .OFRF_REF3_PRE_COSTE = NzDouble(rs("OFRF_REF3_PRE_COSTE"))
            .OFRF_REF3_CE_COSTE = NzDouble(rs("OFRF_REF3_CE_COSTE"))
            .OFRF_REF3_CS_COSTE = NzDouble(rs("OFRF_REF3_CS_COSTE"))
            .OFRF_REF3_IMPORTE = NzDouble(rs("OFRF_REF3_IMPORTE"))
            .OFRF_REF3_DESCRIPCION = NzVBA(rs("OFRF_REF3_DESCRIPCION"))
            .OFRF_REF3_OBSERVA = NzVBA(rs("OFRF_REF3_OBSERVA"))
            .OFRF_REF3_ADVERTENCIA = NzVBA(rs("OFRF_REF3_ADVERTENCIA"))
            .OFRF_REF3_CANTIDAD = NzLong(rs("OFRF_REF3_CANTIDAD"))
            
            ' ========== REFRIGERADOR 4 ==========
            .REF4_ID = NzLong(rs("REF4_ID"))
            .OFRF_REF4_TIPO = NzLong(rs("OFRF_REF4_TIPO"))
            .OFRF_REF4_PN = NzLong(rs("OFRF_REF4_PN"))
            .OFRF_REF4_CE = NzLong(rs("OFRF_REF4_CE"))
            .OFRF_REF4_CS = NzLong(rs("OFRF_REF4_CS"))
            .OFRF_REF4_PRE_COSTE = NzDouble(rs("OFRF_REF4_PRE_COSTE"))
            .OFRF_REF4_CE_COSTE = NzDouble(rs("OFRF_REF4_CE_COSTE"))
            .OFRF_REF4_CS_COSTE = NzDouble(rs("OFRF_REF4_CS_COSTE"))
            .OFRF_REF4_IMPORTE = NzDouble(rs("OFRF_REF4_IMPORTE"))
            .OFRF_REF4_DESCRIPCION = NzVBA(rs("OFRF_REF4_DESCRIPCION"))
            .OFRF_REF4_OBSERVA = NzVBA(rs("OFRF_REF4_OBSERVA"))
            .OFRF_REF4_ADVERTENCIA = NzVBA(rs("OFRF_REF4_ADVERTENCIA"))
            .OFRF_REF4_CANTIDAD = NzLong(rs("OFRF_REF4_CANTIDAD"))
            
            ' ========== REFRIGERADOR 5 ==========
            .REF5_ID = NzLong(rs("REF5_ID"))
            .OFRF_REF5_TIPO = NzLong(rs("OFRF_REF5_TIPO"))
            .OFRF_REF5_PN = NzLong(rs("OFRF_REF5_PN"))
            .OFRF_REF5_CE = NzLong(rs("OFRF_REF5_CE"))
            .OFRF_REF5_CS = NzLong(rs("OFRF_REF5_CS"))
            .OFRF_REF5_PRE_COSTE = NzDouble(rs("OFRF_REF5_PRE_COSTE"))
            .OFRF_REF5_CE_COSTE = NzDouble(rs("OFRF_REF5_CE_COSTE"))
            .OFRF_REF5_CS_COSTE = NzDouble(rs("OFRF_REF5_CS_COSTE"))
            .OFRF_REF5_IMPORTE = NzDouble(rs("OFRF_REF5_IMPORTE"))
            .OFRF_REF5_DESCRIPCION = NzVBA(rs("OFRF_REF5_DESCRIPCION"))
            .OFRF_REF5_OBSERVA = NzVBA(rs("OFRF_REF5_OBSERVA"))
            .OFRF_REF5_ADVERTENCIA = NzVBA(rs("OFRF_REF5_ADVERTENCIA"))
            .OFRF_REF5_CANTIDAD = NzLong(rs("OFRF_REF5_CANTIDAD"))
            
            ' ========== REFRIGERADOR 6 ==========
            .REF6_ID = NzLong(rs("REF6_ID"))
            .OFRF_REF6_TIPO = NzLong(rs("OFRF_REF6_TIPO"))
            .OFRF_REF6_PN = NzLong(rs("OFRF_REF6_PN"))
            .OFRF_REF6_CE = NzLong(rs("OFRF_REF6_CE"))
            .OFRF_REF6_CS = NzLong(rs("OFRF_REF6_CS"))
            .OFRF_REF6_PRE_COSTE = NzDouble(rs("OFRF_REF6_PRE_COSTE"))
            .OFRF_REF6_CE_COSTE = NzDouble(rs("OFRF_REF6_CE_COSTE"))
            .OFRF_REF6_CS_COSTE = NzDouble(rs("OFRF_REF6_CS_COSTE"))
            .OFRF_REF6_IMPORTE = NzDouble(rs("OFRF_REF6_IMPORTE"))
            .OFRF_REF6_DESCRIPCION = NzVBA(rs("OFRF_REF6_DESCRIPCION"))
            .OFRF_REF6_OBSERVA = NzVBA(rs("OFRF_REF6_OBSERVA"))
            .OFRF_REF6_ADVERTENCIA = NzVBA(rs("OFRF_REF6_ADVERTENCIA"))
            .OFRF_REF6_CANTIDAD = NzLong(rs("OFRF_REF6_CANTIDAD"))
            
            ' ========== OPCIONES ==========
            .OFRF_EXTRA_SN = NzVBA(rs("OFRF_EXTRA_SN"))
            .OFRF_RADIO_SN = NzVBA(rs("OFRF_RADIO_SN"))
            .OFRF_ASME_SN = NzVBA(rs("OFRF_ASME_SN"))
            .OFRF_SELLOU_SN = NzVBA(rs("OFRF_SELLOU_SN"))
            .OFRF_SELLOU_IMPORTE = NzLong(rs("OFRF_SELLOU_IMPORTE"))
            
            ' ========== VÁLVULA SEGURIDAD ==========
            .VASE_ID = NzLong(rs("VASE_ID"))
            .OFRF_VASE_PRE_COSTE = NzDouble(rs("OFRF_VASE_PRE_COSTE"))
            .OFRF_VASE_IMPORTE = NzDouble(rs("OFRF_VASE_IMPORTE"))
            .OFRF_VASE_DESCRIPCION = NzVBA(rs("OFRF_VASE_DESCRIPCION"))
            .OFRF_VASE_OBSERVA = NzVBA(rs("OFRF_VASE_OBSERVA"))
            .OFRF_VASE_ADVERTENCIA = NzVBA(rs("OFRF_VASE_ADVERTENCIA"))
            .OFRF_VASE_CANTIDAD = NzLong(rs("OFRF_VASE_CANTIDAD"))
            
            .ImporteCalculado = False
        End With
        Oferta.Refrigeradores = rf
    End If
    
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "CargarRefrigeradores", "Error: " & Err.Description
End Sub

Private Sub CargarAccesorios(ByVal Oferta As clsOferta, ByVal OferID As String)
                                                                           
    On Error GoTo ErrHandler
    
    Dim rs As Object
    Dim ac As tOfertasAccesorios
    
    Set rs = m_Ctx.AbrirRecordset( _
        "SELECT * FROM OfertasAccesorios WHERE OFER_ID={" & OferID & "}")
    
    If Not rs.EOF Then
        With ac
            .OFAC_ID = NormalizarGUID(rs("OFAC_ID"))
            .OFER_ID = NormalizarGUID(rs("OFER_ID"))
            
            ' ========== MOTOR ==========
            .MOTO_ID = NzLong(rs("MOTO_ID"))
            .OFAC_MOTO_PRE_COSTE = NzDouble(rs("OFAC_MOTO_PRE_COSTE"))
            .OFAC_MOTO_IMPORTE = NzDouble(rs("OFAC_MOTO_IMPORTE"))
            .OFAC_MOTO_DESCRIPCION = NzVBA(rs("OFAC_MOTO_DESCRIPCION"))
            .OFAC_MOTO_OBSERVA = NzVBA(rs("OFAC_MOTO_OBSERVA"))
            .OFAC_MOTO_ADVERTENCIA = NzVBA(rs("OFAC_MOTO_ADVERTENCIA"))
            .OFAC_MOTO_TIPO = NzLong(rs("OFAC_MOTO_TIPO"))
            .OFAC_MOTO_ESP = NzVBA(rs("OFAC_MOTO_ESP"))
            .OFAC_MOTO_ESP_IMPORTE = NzDouble(rs("OFAC_MOTO_ESP_IMPORTE"))
            
            ' ========== ARRANCADOR ==========
            .ARRA_ID = NzLong(rs("ARRA_ID"))
            .OFAC_ARRA_PRE_COSTE = NzDouble(rs("OFAC_ARRA_PRE_COSTE"))
            .OFAC_ARRA_IMPORTE = NzDouble(rs("OFAC_ARRA_IMPORTE"))
            .OFAC_ARRA_DESCRIPCION = NzVBA(rs("OFAC_ARRA_DESCRIPCION"))
            .OFAC_ARRA_OBSERVA = NzVBA(rs("OFAC_ARRA_OBSERVA"))
            .OFAC_ARRA_ADVERTENCIA = NzVBA(rs("OFAC_ARRA_ADVERTENCIA"))
            .OFAC_ARRA_CANTIDAD = NzLong(rs("OFAC_ARRA_CANTIDAD"))
            .OFAC_ARRA_TIPO = NzLong(rs("OFAC_ARRA_TIPO"))
            
            ' ========== CAJA LOCAL ==========
            .CALO_ID = NzLong(rs("CALO_ID"))
            .OFAC_CALO_PRE_COSTE = NzDouble(rs("OFAC_CALO_PRE_COSTE"))
            .OFAC_CALO_IMPORTE = NzDouble(rs("OFAC_CALO_IMPORTE"))
            .OFAC_CALO_DESCRIPCION = NzVBA(rs("OFAC_CALO_DESCRIPCION"))
            .OFAC_CALO_OBSERVA = NzVBA(rs("OFAC_CALO_OBSERVA"))
            .OFAC_CALO_ADVERTENCIA = NzVBA(rs("OFAC_CALO_ADVERTENCIA"))
            
            ' ========== FILTRO ==========
            .FILT_ID = NzLong(rs("FILT_ID"))
            .OFAC_FILT_PRE_COSTE = NzDouble(rs("OFAC_FILT_PRE_COSTE"))
            .OFAC_FILT_IMPORTE = NzDouble(rs("OFAC_FILT_IMPORTE"))
            .OFAC_FILT_DESCRIPCION = NzVBA(rs("OFAC_FILT_DESCRIPCION"))
            .OFAC_FILT_OBSERVA = NzVBA(rs("OFAC_FILT_OBSERVA"))
            .OFAC_FILT_ADVERTENCIA = NzVBA(rs("OFAC_FILT_ADVERTENCIA"))
            
            ' ========== AERO ==========
            .AERO_ID = NzLong(rs("AERO_ID"))
            .OFAC_AERO_PRE_COSTE = NzDouble(rs("OFAC_AERO_PRE_COSTE"))
            .OFAC_AERO_IMPORTE = NzDouble(rs("OFAC_AERO_IMPORTE"))
            .OFAC_AERO_DESCRIPCION = NzVBA(rs("OFAC_AERO_DESCRIPCION"))
            .OFAC_AERO_OBSERVA = NzVBA(rs("OFAC_AERO_OBSERVA"))
            .OFAC_AERO_ADVERTENCIA = NzVBA(rs("OFAC_AERO_ADVERTENCIA"))
            
            ' ========== LLAVE ENTRADA ==========
            .LLEN_ID = NzLong(rs("LLEN_ID"))
            .OFAC_LLEN_PRE_COSTE = NzDouble(rs("OFAC_LLEN_PRE_COSTE"))
            .OFAC_LLEN_IMPORTE = NzDouble(rs("OFAC_LLEN_IMPORTE"))
            .OFAC_LLEN_DESCRIPCION = NzVBA(rs("OFAC_LLEN_DESCRIPCION"))
            .OFAC_LLEN_OBSERVA = NzVBA(rs("OFAC_LLEN_OBSERVA"))
            .OFAC_LLEN_ADVERTENCIA = NzVBA(rs("OFAC_LLEN_ADVERTENCIA"))
            
            ' ========== LLAVE SALIDA ==========
            .LLSA_ID = NzLong(rs("LLSA_ID"))
            .OFAC_LLSA_PRE_COSTE = NzDouble(rs("OFAC_LLSA_PRE_COSTE"))
            .OFAC_LLSA_IMPORTE = NzDouble(rs("OFAC_LLSA_IMPORTE"))
            .OFAC_LLSA_DESCRIPCION = NzVBA(rs("OFAC_LLSA_DESCRIPCION"))
            .OFAC_LLSA_OBSERVA = NzVBA(rs("OFAC_LLSA_OBSERVA"))
            .OFAC_LLSA_ADVERTENCIA = NzVBA(rs("OFAC_LLSA_ADVERTENCIA"))
            
            ' ========== GRUPO ENGRASE ==========
            .GREN_ID = NzLong(rs("GREN_ID"))
            .OFAC_GREN_PRE_COSTE = NzDouble(rs("OFAC_GREN_PRE_COSTE"))
            .OFAC_GREN_IMPORTE = NzDouble(rs("OFAC_GREN_IMPORTE"))
            .OFAC_GREN_DESCRIPCION = NzVBA(rs("OFAC_GREN_DESCRIPCION"))
            .OFAC_GREN_OBSERVA = NzVBA(rs("OFAC_GREN_OBSERVA"))
            .OFAC_GREN_ADVERTENCIA = NzVBA(rs("OFAC_GREN_ADVERTENCIA"))
            
            .ImporteCalculado = False
        End With
        Oferta.Accesorios = ac
    End If
    
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "CargarAccesorios", "Error: " & Err.Description
End Sub

Private Sub CargarInstrumentacion(ByVal Oferta As clsOferta, ByVal OferID As String)
                                                                           
    On Error GoTo ErrHandler
    
    Dim rs As Object
    Dim inst As tOfertasInstrumentacion
    
    Set rs = m_Ctx.AbrirRecordset( _
        "SELECT * FROM OfertasInstrumentacion WHERE OFER_ID={" & OferID & "}")
    
    If Not rs.EOF Then
        With inst
            .OFIN_ID = NormalizarGUID(rs("OFIN_ID"))
            .OFER_ID = NormalizarGUID(rs("OFER_ID"))
            
            ' ========== INSTRUMENTACIÓN BASE ==========
            .INTR_ID = NzLong(rs("INTR_ID"))
            .OFIN_INTR_DESCRIPCION = NzVBA(rs("OFIN_INTR_DESCRIPCION"))
            .OFIN_INTR_OBSERVA = NzVBA(rs("OFIN_INTR_OBSERVA"))
            .OFIN_INTR_ADVERTENCIA = NzVBA(rs("OFIN_INTR_ADVERTENCIA"))
            
            ' ========== TRANSMISOR TEMPERATURA ==========
            .TRAT_ID = NzLong(rs("TRAT_ID"))
            .OFIN_TRAT_PRE_COSTE = NzDouble(rs("OFIN_TRAT_PRE_COSTE"))
            .OFIN_TRAT_IMPORTE = NzDouble(rs("OFIN_TRAT_IMPORTE"))
            .OFIN_TRAT_DESCRIPCION = NzVBA(rs("OFIN_TRAT_DESCRIPCION"))
            .OFIN_TRAT_OBSERVA = NzVBA(rs("OFIN_TRAT_OBSERVA"))
            .OFIN_TRAT_ADVERTENCIA = NzVBA(rs("OFIN_TRAT_ADVERTENCIA"))
            .OFIN_TRAT_CANTIDAD = NzLong(rs("OFIN_TRAT_CANTIDAD"))
            
            ' ========== TRANSMISOR PRESIÓN ==========
            .TRAP_ID = NzLong(rs("TRAP_ID"))
            .OFIN_TRAP_PRE_COSTE = NzDouble(rs("OFIN_TRAP_PRE_COSTE"))
            .OFIN_TRAP_IMPORTE = NzDouble(rs("OFIN_TRAP_IMPORTE"))
            .OFIN_TRAP_DESCRIPCION = NzVBA(rs("OFIN_TRAP_DESCRIPCION"))
            .OFIN_TRAP_OBSERVA = NzVBA(rs("OFIN_TRAP_OBSERVA"))
            .OFIN_TRAP_ADVERTENCIA = NzVBA(rs("OFIN_TRAP_ADVERTENCIA"))
            .OFIN_TRAP_CANTIDAD = NzLong(rs("OFIN_TRAP_CANTIDAD"))
            
            ' ========== TERMÓMETRO ==========
            .TERM_ID = NzLong(rs("TERM_ID"))
            .OFIN_TERM_PRE_COSTE = NzDouble(rs("OFIN_TERM_PRE_COSTE"))
            .OFIN_TERM_IMPORTE = NzDouble(rs("OFIN_TERM_IMPORTE"))
            .OFIN_TERM_DESCRIPCION = NzVBA(rs("OFIN_TERM_DESCRIPCION"))
            .OFIN_TERM_OBSERVA = NzVBA(rs("OFIN_TERM_OBSERVA"))
            .OFIN_TERM_ADVERTENCIA = NzVBA(rs("OFIN_TERM_ADVERTENCIA"))
            .OFIN_TERM_CANTIDAD = NzLong(rs("OFIN_TERM_CANTIDAD"))
            
            ' ========== MANÓMETRO ==========
            .MANO_ID = NzLong(rs("MANO_ID"))
            .OFIN_MANO_PRE_COSTE = NzDouble(rs("OFIN_MANO_PRE_COSTE"))
            .OFIN_MANO_IMPORTE = NzDouble(rs("OFIN_MANO_IMPORTE"))
            .OFIN_MANO_DESCRIPCION = NzVBA(rs("OFIN_MANO_DESCRIPCION"))
            .OFIN_MANO_OBSERVA = NzVBA(rs("OFIN_MANO_OBSERVA"))
            .OFIN_MANO_ADVERTENCIA = NzVBA(rs("OFIN_MANO_ADVERTENCIA"))
            .OFIN_MANO_CANTIDAD = NzLong(rs("OFIN_MANO_CANTIDAD"))
            
            ' ========== ELECTROVÁLVULA REGULACIÓN ==========
            .ELER_ID = NzLong(rs("ELER_ID"))
            .OFIN_ELER_PRE_COSTE = NzDouble(rs("OFIN_ELER_PRE_COSTE"))
            .OFIN_ELER_IMPORTE = NzDouble(rs("OFIN_ELER_IMPORTE"))
            .OFIN_ELER_DESCRIPCION = NzVBA(rs("OFIN_ELER_DESCRIPCION"))
            .OFIN_ELER_OBSERVA = NzVBA(rs("OFIN_ELER_OBSERVA"))
            .OFIN_ELER_ADVERTENCIA = NzVBA(rs("OFIN_ELER_ADVERTENCIA"))
            .OFIN_ELER_CANTIDAD = NzLong(rs("OFIN_ELER_CANTIDAD"))
            
            ' ========== SENSOR CAÍDA VÁSTAGO ==========
            .SECV_ID = NzLong(rs("SECV_ID"))
            .OFIN_SECV_PRE_COSTE = NzDouble(rs("OFIN_SECV_PRE_COSTE"))
            .OFIN_SECV_IMPORTE = NzDouble(rs("OFIN_SECV_IMPORTE"))
            .OFIN_SECV_DESCRIPCION = NzVBA(rs("OFIN_SECV_DESCRIPCION"))
            .OFIN_SECV_OBSERVA = NzVBA(rs("OFIN_SECV_OBSERVA"))
            .OFIN_SECV_ADVERTENCIA = NzVBA(rs("OFIN_SECV_ADVERTENCIA"))
            .OFIN_SECV_CANTIDAD = NzLong(rs("OFIN_SECV_CANTIDAD"))
            
            ' ========== INTERRUPTOR VIBRACIÓN ==========
            .INTV_ID = NzLong(rs("INTV_ID"))
            .OFIN_INTV_PRE_COSTE = NzDouble(rs("OFIN_INTV_PRE_COSTE"))
            .OFIN_INTV_IMPORTE = NzDouble(rs("OFIN_INTV_IMPORTE"))
            .OFIN_INTV_DESCRIPCION = NzVBA(rs("OFIN_INTV_DESCRIPCION"))
            .OFIN_INTV_OBSERVA = NzVBA(rs("OFIN_INTV_OBSERVA"))
            .OFIN_INTV_ADVERTENCIA = NzVBA(rs("OFIN_INTV_ADVERTENCIA"))
            .OFIN_INTV_CANTIDAD = NzLong(rs("OFIN_INTV_CANTIDAD"))
            
            ' ========== INTERRUPTOR NIVEL ACEITE ==========
            .INTA_ID = NzLong(rs("INTA_ID"))
            .OFIN_INTA_PRE_COSTE = NzDouble(rs("OFIN_INTA_PRE_COSTE"))
            .OFIN_INTA_IMPORTE = NzDouble(rs("OFIN_INTA_IMPORTE"))
            .OFIN_INTA_DESCRIPCION = NzVBA(rs("OFIN_INTA_DESCRIPCION"))
            .OFIN_INTA_OBSERVA = NzVBA(rs("OFIN_INTA_OBSERVA"))
            .OFIN_INTA_ADVERTENCIA = NzVBA(rs("OFIN_INTA_ADVERTENCIA"))
            .OFIN_INTA_CANTIDAD = NzLong(rs("OFIN_INTA_CANTIDAD"))
            
            ' ========== NIVEL CONDENSADOS ==========
            .NIVC_ID = NzLong(rs("NIVC_ID"))
            .OFIN_NIVC_PRE_COSTE = NzDouble(rs("OFIN_NIVC_PRE_COSTE"))
            .OFIN_NIVC_IMPORTE = NzDouble(rs("OFIN_NIVC_IMPORTE"))
            .OFIN_NIVC_DESCRIPCION = NzVBA(rs("OFIN_NIVC_DESCRIPCION"))
            .OFIN_NIVC_OBSERVA = NzVBA(rs("OFIN_NIVC_OBSERVA"))
            .OFIN_NIVC_ADVERTENCIA = NzVBA(rs("OFIN_NIVC_ADVERTENCIA"))
            .OFIN_NIVC_CANTIDAD = NzLong(rs("OFIN_NIVC_CANTIDAD"))
            
            ' ========== VÁLVULA TERMOSTÁTICA ==========
            .VATE_ID = NzLong(rs("VATE_ID"))
            .OFIN_VATE_PRE_COSTE = NzDouble(rs("OFIN_VATE_PRE_COSTE"))
            .OFIN_VATE_IMPORTE = NzDouble(rs("OFIN_VATE_IMPORTE"))
            .OFIN_VATE_DESCRIPCION = NzVBA(rs("OFIN_VATE_DESCRIPCION"))
            .OFIN_VATE_OBSERVA = NzVBA(rs("OFIN_VATE_OBSERVA"))
            .OFIN_VATE_ADVERTENCIA = NzVBA(rs("OFIN_VATE_ADVERTENCIA"))
            .OFIN_VATE_CANTIDAD = NzLong(rs("OFIN_VATE_CANTIDAD"))
            
            ' ========== INSTRUMENTACIÓN LIBRE (7 LÍNEAS) ==========
            .OFIN_INS1_TEXTO = NzVBA(rs("OFIN_INS1_TEXTO"))
            .OFIN_INS1_CANTIDAD = NzInteger(rs("OFIN_INS1_CANTIDAD"))
            .OFIN_INS1_IMPORTE = NzDouble(rs("OFIN_INS1_IMPORTE"))
            
            .OFIN_INS2_TEXTO = NzVBA(rs("OFIN_INS2_TEXTO"))
            .OFIN_INS2_CANTIDAD = NzInteger(rs("OFIN_INS2_CANTIDAD"))
            .OFIN_INS2_IMPORTE = NzDouble(rs("OFIN_INS2_IMPORTE"))
            
            .OFIN_INS3_TEXTO = NzVBA(rs("OFIN_INS3_TEXTO"))
            .OFIN_INS3_CANTIDAD = NzInteger(rs("OFIN_INS3_CANTIDAD"))
            .OFIN_INS3_IMPORTE = NzDouble(rs("OFIN_INS3_IMPORTE"))
            
            .OFIN_INS4_TEXTO = NzVBA(rs("OFIN_INS4_TEXTO"))
            .OFIN_INS4_CANTIDAD = NzInteger(rs("OFIN_INS4_CANTIDAD"))
            .OFIN_INS4_IMPORTE = NzDouble(rs("OFIN_INS4_IMPORTE"))
            
            .OFIN_INS5_TEXTO = NzVBA(rs("OFIN_INS5_TEXTO"))
            .OFIN_INS5_CANTIDAD = NzInteger(rs("OFIN_INS5_CANTIDAD"))
            .OFIN_INS5_IMPORTE = NzDouble(rs("OFIN_INS5_IMPORTE"))
            
            .OFIN_INS6_TEXTO = NzVBA(rs("OFIN_INS6_TEXTO"))
            .OFIN_INS6_CANTIDAD = NzInteger(rs("OFIN_INS6_CANTIDAD"))
            .OFIN_INS6_IMPORTE = NzDouble(rs("OFIN_INS6_IMPORTE"))
            
            .OFIN_INS7_TEXTO = NzVBA(rs("OFIN_INS7_TEXTO"))
            .OFIN_INS7_CANTIDAD = NzInteger(rs("OFIN_INS7_CANTIDAD"))
            .OFIN_INS7_IMPORTE = NzDouble(rs("OFIN_INS7_IMPORTE"))
            
            .ImporteCalculado = False
        End With
        Oferta.Instrumentacion = inst
    End If
    
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "CargarInstrumentacion", "Error: " & Err.Description
End Sub

Private Sub CargarExtras(ByVal Oferta As clsOferta, ByVal OferID As String)
                                                                           
    On Error GoTo ErrHandler
    
    Dim rs As Object
    Dim ex As tOfertasExtras
    
    Set rs = m_Ctx.AbrirRecordset( _
        "SELECT * FROM OfertasExtras WHERE OFER_ID={" & OferID & "}")
    
    If Not rs.EOF Then
        With ex
            .OFEX_ID = NormalizarGUID(rs("OFEX_ID"))
            .OFER_ID = NormalizarGUID(rs("OFER_ID"))
            
            ' ========== EMBALAJE ==========
            .EMBA_ID = NzLong(rs("EMBA_ID"))
            .OFEX_EMBA_PRE_COSTE = NzDouble(rs("OFEX_EMBA_PRE_COSTE"))
            .OFEX_EMBA_IMPORTE = NzDouble(rs("OFEX_EMBA_IMPORTE"))
            .OFEX_EMBA_DESCRIPCION = NzVBA(rs("OFEX_EMBA_DESCRIPCION"))
            .OFEX_EMBA_OBSERVA = NzVBA(rs("OFEX_EMBA_OBSERVA"))
            .OFEX_EMBA_ADVERTENCIA = NzVBA(rs("OFEX_EMBA_ADVERTENCIA"))
            
            ' ========== TRANSPORTE ==========
            .TRAN_ID = NzLong(rs("TRAN_ID"))
            .OFEX_TRAN_PRE_COSTE = NzDouble(rs("OFEX_TRAN_PRE_COSTE"))
            .OFEX_TRAN_IMPORTE = NzDouble(rs("OFEX_TRAN_IMPORTE"))
            .OFEX_TRAN_DESCRIPCION = NzVBA(rs("OFEX_TRAN_DESCRIPCION"))
            .OFEX_TRAN_OBSERVA = NzVBA(rs("OFEX_TRAN_OBSERVA"))
            .OFEX_TRAN_ADVERTENCIA = NzVBA(rs("OFEX_TRAN_ADVERTENCIA"))
            
            ' ========== PUESTA EN MARCHA ==========
            .PUMA_ID = NzLong(rs("PUMA_ID"))
            .OFEX_PUMA_PRE_COSTE = NzDouble(rs("OFEX_PUMA_PRE_COSTE"))
            .OFEX_PUMA_IMPORTE = NzDouble(rs("OFEX_PUMA_IMPORTE"))
            .OFEX_PUMA_DESCRIPCION = NzVBA(rs("OFEX_PUMA_DESCRIPCION"))
            .OFEX_PUMA_OBSERVA = NzVBA(rs("OFEX_PUMA_OBSERVA"))
            .OFEX_PUMA_ADVERTENCIA = NzVBA(rs("OFEX_PUMA_ADVERTENCIA"))
            
            ' ========== PRUEBAS ASISTENCIA 1 ==========
            .PRA1_ID = NzLong(rs("PRA1_ID"))
            .OFEX_PRA1_PRE_COSTE = NzDouble(rs("OFEX_PRA1_PRE_COSTE"))
            .OFEX_PRA1_IMPORTE = NzDouble(rs("OFEX_PRA1_IMPORTE"))
            .OFEX_PRA1_DESCRIPCION = NzVBA(rs("OFEX_PRA1_DESCRIPCION"))
            .OFEX_PRA1_OBSERVA = NzVBA(rs("OFEX_PRA1_OBSERVA"))
            .OFEX_PRA1_ADVERTENCIA = NzVBA(rs("OFEX_PRA1_ADVERTENCIA"))
            
            ' ========== PRUEBAS ASISTENCIA 2 ==========
            .PRA2_ID = NzLong(rs("PRA2_ID"))
            .OFEX_PRA2_PRE_COSTE = NzDouble(rs("OFEX_PRA2_PRE_COSTE"))
            .OFEX_PRA2_IMPORTE = NzDouble(rs("OFEX_PRA2_IMPORTE"))
            .OFEX_PRA2_DESCRIPCION = NzVBA(rs("OFEX_PRA2_DESCRIPCION"))
            .OFEX_PRA2_OBSERVA = NzVBA(rs("OFEX_PRA2_OBSERVA"))
            .OFEX_PRA2_ADVERTENCIA = NzVBA(rs("OFEX_PRA2_ADVERTENCIA"))
            
            ' ========== PRUEBAS ASISTENCIA 3 ==========
            .PRA3_ID = NzLong(rs("PRA3_ID"))
            .OFEX_PRA3_PRE_COSTE = NzDouble(rs("OFEX_PRA3_PRE_COSTE"))
            .OFEX_PRA3_IMPORTE = NzDouble(rs("OFEX_PRA3_IMPORTE"))
            .OFEX_PRA3_DESCRIPCION = NzVBA(rs("OFEX_PRA3_DESCRIPCION"))
            .OFEX_PRA3_OBSERVA = NzVBA(rs("OFEX_PRA3_OBSERVA"))
            .OFEX_PRA3_ADVERTENCIA = NzVBA(rs("OFEX_PRA3_ADVERTENCIA"))
            
            ' ========== PRUEBAS ASISTENCIA 4 ==========
            .PRA4_ID = NzLong(rs("PRA4_ID"))
            .OFEX_PRA4_PRE_COSTE = NzDouble(rs("OFEX_PRA4_PRE_COSTE"))
            .OFEX_PRA4_IMPORTE = NzDouble(rs("OFEX_PRA4_IMPORTE"))
            .OFEX_PRA4_DESCRIPCION = NzVBA(rs("OFEX_PRA4_DESCRIPCION"))
            .OFEX_PRA4_OBSERVA = NzVBA(rs("OFEX_PRA4_OBSERVA"))
            .OFEX_PRA4_ADVERTENCIA = NzVBA(rs("OFEX_PRA4_ADVERTENCIA"))
            
            .ImporteCalculado = False
        End With
        Oferta.Extras = ex
    End If
    
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "CargarExtras", "Error: " & Err.Description
End Sub

Private Sub CargarOtros(ByVal Oferta As clsOferta, ByVal OferID As String)
    ' IMPLEMENTACIÓN COMPLETA - NO MODIFICAR
    On Error GoTo ErrHandler

    Dim rs As Object

    Set rs = m_Ctx.AbrirRecordset( _
        "SELECT * FROM OfertasOtros WHERE OFER_ID={" & OferID & "} ORDER BY OFOT_LINEA")

    Dim it As New clsOfertaOtro
    Do While Not rs.EOF
        Set it = New clsOfertaOtro
        it.OFOT_ID = NormalizarGUID(rs("OFOT_ID"))
        it.OFER_ID = NormalizarGUID(rs("OFER_ID"))
        it.OFOT_LINEA = NzLong(rs("OFOT_LINEA"))
        it.OFOT_DESCRIPCION = NzVBA(rs("OFOT_DESCRIPCION"))
        it.OFOT_PRE_COSTE = NzDouble(rs("OFOT_PRE_COSTE"))
        Oferta.AgregarOtro it
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "CargarOtros", "Error: " & Err.Description
End Sub

'=========================================================
' FUNCIONES AUXILIARES - NO MODIFICAR
'=========================================================

Private Function NormalizarGUID(ByVal v As Variant) As String
    NormalizarGUID = Replace(Replace(v, "{", ""), "}", "")
End Function

Private Function NzVBA(ByVal v As Variant) As String
    If IsNull(v) Then NzVBA = "" Else NzVBA = CStr(v)
End Function

Private Function NzLong(ByVal v As Variant) As Long
    If IsNull(v) Then NzLong = 0 Else NzLong = CLng(v)
End Function

Private Function NzDouble(ByVal v As Variant) As Double
    If IsNull(v) Then NzDouble = 0 Else NzDouble = CDbl(v)
End Function

Private Function NzDate(ByVal v As Variant) As Date
    If IsNull(v) Then NzDate = 0 Else NzDate = CDate(v)
End Function

Private Function NzInteger(ByVal v As Variant) As Integer
    If IsNull(v) Then NzInteger = 0 Else NzInteger = CInt(v)
End Function

Private Function isGUID(ByVal strGUID As Variant) As Boolean
    On Error GoTo ErrHandler
    If IsNull(strGUID) Or IsEmpty(strGUID) Then
        isGUID = False
        Exit Function
    End If
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "[0-9A-Fa-f]{8}-(?:[0-9A-Fa-f]{4}-){3}[0-9A-Fa-f]{12}"
    isGUID = regEx.Test(CStr(strGUID))
    Set regEx = Nothing
    Exit Function
ErrHandler:
    isGUID = False
End Function
