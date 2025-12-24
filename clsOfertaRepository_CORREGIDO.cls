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

'
' TODO: Copiar las versiones completas desde clsOfertaRepository_SECCIONES.txt
'       de los siguientes métodos:
'       - CargarCabezal (6 cilindros)
'       - CargarCalderines (6 etapas)
'       - CargarRefrigeradores (6 refrigeradores)
'       - CargarAccesorios
'       - CargarInstrumentacion
'       - CargarExtras
'

Private Sub CargarDatosGenerales(ByVal Oferta As clsOferta, ByVal OferID As String)
    ' IMPLEMENTACIÓN COMPLETA - NO MODIFICAR
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
        .PAIS_ID = NzLong(rs("PAIS_ID"))
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

        ' Resolver FKs de datos generales
        If .PAIS_ID > 0 Then
            Dim pais As tCatalogPaises
            pais = m_Catalogos.ObtenerPaises(.PAIS_ID)
            .pais_nombre = pais.pais_nombre
        End If

        If .COME_ID > 0 Then
            Dim come As tCatalogComerciales
            come = m_Catalogos.ObtenerComerciales(.COME_ID)
            .COME_NOMBRE = come.COME_NOMBRE
        End If

        If .OFTN_ID > 0 Then
            Dim oftn As tCatalogOfertantes
            oftn = m_Catalogos.ObtenerOfertantes(.OFTN_ID)
            .OFTN_NOMBRE = oftn.OFTN_NOMBRE
        End If

        If .GASE_ID > 0 Then
            Dim gase As tCatalogGases
            gase = m_Catalogos.ObtenerGases(.GASE_ID)
            .GASE_DENOMINACION = gase.GASE_DENOMINACION
        End If
    End With

    Oferta.DatosGenerales = dg
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "CargarDatosGenerales", "Error: " & Err.Description
End Sub

Private Sub CargarManoObra(ByVal Oferta As clsOferta, ByVal OferID As String)
    ' IMPLEMENTACIÓN COMPLETA - NO MODIFICAR
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
    ' IMPLEMENTACIÓN COMPLETA - NO MODIFICAR
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
    ' IMPLEMENTACIÓN COMPLETA - NO MODIFICAR
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

            ' Resolver FKs de opciones
            If .VAPR_ID > 0 Then
                Dim vapr As tCatalogValvulasPresion
                vapr = m_Catalogos.ObtenerValvulasPresion(.VAPR_ID)
                If .OFOP_VAPR_DESCRIPCION = "" Then
                    .OFOP_VAPR_DESCRIPCION = vapr.VAPR_MODELO & " - " & vapr.VAPR_DESCRIPCION
                End If
                If .OFOP_VAPR_PRE_COSTE = 0 Then
                    .OFOP_VAPR_PRE_COSTE = vapr.VAPR_PRE_COSTE
                End If
            End If

            If .ENCI_ID > 0 Then
                Dim enci As tCatalogEngraseCilindros
                enci = m_Catalogos.ObtenerEngraseCilindros(.ENCI_ID)
                If .OFOP_ENCI_DESCRIPCION = "" Then
                    .OFOP_ENCI_DESCRIPCION = enci.ENCI_MODELO & " - " & enci.ENCI_DESCRIPCION
                End If
                If .OFOP_ENCI_PRE_COSTE = 0 Then
                    .OFOP_ENCI_PRE_COSTE = enci.ENCI_PRE_COSTE
                End If
            End If

            If .PURG_ID > 0 Then
                Dim purg As tCatalogPurgadores
                purg = m_Catalogos.ObtenerPurgadores(.PURG_ID)
                If .OFOP_PURG_DESCRIPCION = "" Then
                    .OFOP_PURG_DESCRIPCION = purg.PURG_MODELO & " - " & purg.PURG_DESCRIPCION
                End If
                If .OFOP_PURG_PRE_COSTE = 0 Then
                    .OFOP_PURG_PRE_COSTE = purg.PURG_PRE_COSTE
                End If
            End If

            If .RESC_ID > 0 Then
                Dim resc As tCatalogResistenciasCalefaccion
                resc = m_Catalogos.ObtenerResistenciasCalefaccion(.RESC_ID)
                If .OFOP_RESC_DESCRIPCION = "" Then
                    .OFOP_RESC_DESCRIPCION = resc.RESC_MODELO & " - " & resc.RESC_DESCRIPCION
                End If
                If .OFOP_RESC_PRE_COSTE = 0 Then
                    .OFOP_RESC_PRE_COSTE = resc.RESC_PRE_COSTE
                End If
            End If

            If .VARG_ID > 0 Then
                Dim varg As tCatalogValvulasReguladoras
                varg = m_Catalogos.ObtenerValvulasReguladoras(.VARG_ID)
                If .OFOP_VARG_DESCRIPCION = "" Then
                    .OFOP_VARG_DESCRIPCION = varg.VARG_MODELO & " - " & varg.VARG_DESCRIPCION
                End If
                If .OFOP_VARG_PRE_COSTE = 0 Then
                    .OFOP_VARG_PRE_COSTE = varg.VARG_PRE_COSTE
                End If
            End If

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

'
' ⚠️ IMPORTANTE: Las siguientes funciones están ABREVIADAS
' Debes copiar las versiones completas desde clsOfertaRepository_SECCIONES.txt
'

Private Sub CargarCabezal(ByVal Oferta As clsOferta, ByVal OferID As String)
    ' TODO: COPIAR VERSIÓN COMPLETA desde clsOfertaRepository_SECCIONES.txt
    ' Esta versión es solo un esqueleto
    On Error GoTo ErrHandler
    Dim rs As Object
    Dim cb As tOfertasCabezal
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM OfertasCabezal WHERE OFER_ID={" & OferID & "}")
    If Not rs.EOF Then
        cb.OFCA_ID = NormalizarGUID(rs("OFCA_ID"))
        cb.OFER_ID = NormalizarGUID(rs("OFER_ID"))
        ' ... FALTA IMPLEMENTAR 6 CILINDROS ...
        Oferta.Cabezal = cb
    End If
    rs.Close
    Exit Sub
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "CargarCabezal", "Error: " & Err.Description
End Sub

Private Sub CargarCalderines(ByVal Oferta As clsOferta, ByVal OferID As String)
    ' TODO: COPIAR VERSIÓN COMPLETA desde clsOfertaRepository_SECCIONES.txt
    On Error GoTo ErrHandler
    Dim rs As Object
    Dim cl As tOfertasCalderines
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM OfertasCalderines WHERE OFER_ID={" & OferID & "}")
    If Not rs.EOF Then
        cl.OFCL_ID = NzLong(rs("OFCL_ID"))
        cl.OFER_ID = NormalizarGUID(rs("OFER_ID"))
        ' ... FALTA IMPLEMENTAR 6 ETAPAS ...
        Oferta.Calderines = cl
    End If
    rs.Close
    Exit Sub
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "CargarCalderines", "Error: " & Err.Description
End Sub

Private Sub CargarRefrigeradores(ByVal Oferta As clsOferta, ByVal OferID As String)
    ' TODO: COPIAR VERSIÓN COMPLETA desde clsOfertaRepository_SECCIONES.txt
    On Error GoTo ErrHandler
    Dim rs As Object
    Dim rf As tOfertasRefrigeradores
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM OfertasRefrigeradores WHERE OFER_ID={" & OferID & "}")
    If Not rs.EOF Then
        rf.OFRF_ID = NormalizarGUID(rs("OFRF_ID"))
        rf.OFER_ID = NormalizarGUID(rs("OFER_ID"))
        ' ... FALTA IMPLEMENTAR 6 REFRIGERADORES ...
        Oferta.Refrigeradores = rf
    End If
    rs.Close
    Exit Sub
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "CargarRefrigeradores", "Error: " & Err.Description
End Sub

Private Sub CargarAccesorios(ByVal Oferta As clsOferta, ByVal OferID As String)
    ' TODO: COPIAR VERSIÓN COMPLETA desde clsOfertaRepository_SECCIONES.txt
    On Error GoTo ErrHandler
    Dim rs As Object
    Dim ac As tOfertasAccesorios
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM OfertasAccesorios WHERE OFER_ID={" & OferID & "}")
    If Not rs.EOF Then
        ac.OFAC_ID = NormalizarGUID(rs("OFAC_ID"))
        ac.OFER_ID = NormalizarGUID(rs("OFER_ID"))
        ' ... FALTA IMPLEMENTAR TODOS LOS ACCESORIOS ...
        Oferta.Accesorios = ac
    End If
    rs.Close
    Exit Sub
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "CargarAccesorios", "Error: " & Err.Description
End Sub

Private Sub CargarInstrumentacion(ByVal Oferta As clsOferta, ByVal OferID As String)
    ' TODO: COPIAR VERSIÓN COMPLETA desde clsOfertaRepository_SECCIONES.txt
    On Error GoTo ErrHandler
    Dim rs As Object
    Dim inst As tOfertasInstrumentacion
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM OfertasInstrumentacion WHERE OFER_ID={" & OferID & "}")
    If Not rs.EOF Then
        inst.OFIN_ID = NormalizarGUID(rs("OFIN_ID"))
        inst.OFER_ID = NormalizarGUID(rs("OFER_ID"))
        ' ... FALTA IMPLEMENTAR TODA LA INSTRUMENTACIÓN ...
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
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM OfertasExtras WHERE OFER_ID={" & OferID & "}")
    If Not rs.EOF Then
        With ex
            .OFEX_ID = NormalizarGUID(rs("OFEX_ID"))
            .OFER_ID = NormalizarGUID(rs("OFER_ID"))
            ' Embalaje
            .EMBA_ID = NzLong(rs("EMBA_ID"))
            .OFEX_EMBA_PRE_COSTE = NzDouble(rs("OFEX_EMBA_PRE_COSTE"))
            .OFEX_EMBA_IMPORTE = NzDouble(rs("OFEX_EMBA_IMPORTE"))
            .OFEX_EMBA_DESCRIPCION = NzVBA(rs("OFEX_EMBA_DESCRIPCION"))
            .OFEX_EMBA_OBSERVA = NzVBA(rs("OFEX_EMBA_OBSERVA"))
            .OFEX_EMBA_ADVERTENCIA = NzVBA(rs("OFEX_EMBA_ADVERTENCIA"))
            ' Transporte
            .TRAN_ID = NzLong(rs("TRAN_ID"))
            .OFEX_TRAN_PRE_COSTE = NzDouble(rs("OFEX_TRAN_PRE_COSTE"))
            .OFEX_TRAN_IMPORTE = NzDouble(rs("OFEX_TRAN_IMPORTE"))
            .OFEX_TRAN_DESCRIPCION = NzVBA(rs("OFEX_TRAN_DESCRIPCION"))
            .OFEX_TRAN_OBSERVA = NzVBA(rs("OFEX_TRAN_OBSERVA"))
            .OFEX_TRAN_ADVERTENCIA = NzVBA(rs("OFEX_TRAN_ADVERTENCIA"))
            ' Puesta en marcha
            .PUMA_ID = NzLong(rs("PUMA_ID"))
            .OFEX_PUMA_PRE_COSTE = NzDouble(rs("OFEX_PUMA_PRE_COSTE"))
            .OFEX_PUMA_IMPORTE = NzDouble(rs("OFEX_PUMA_IMPORTE"))
            .OFEX_PUMA_DESCRIPCION = NzVBA(rs("OFEX_PUMA_DESCRIPCION"))
            .OFEX_PUMA_OBSERVA = NzVBA(rs("OFEX_PUMA_OBSERVA"))
            .OFEX_PUMA_ADVERTENCIA = NzVBA(rs("OFEX_PUMA_ADVERTENCIA"))
            ' Pruebas asistencia 1
            .PRA1_ID = NzLong(rs("PRA1_ID"))
            .OFEX_PRA1_PRE_COSTE = NzDouble(rs("OFEX_PRA1_PRE_COSTE"))
            .OFEX_PRA1_IMPORTE = NzDouble(rs("OFEX_PRA1_IMPORTE"))
            .OFEX_PRA1_DESCRIPCION = NzVBA(rs("OFEX_PRA1_DESCRIPCION"))
            .OFEX_PRA1_OBSERVA = NzVBA(rs("OFEX_PRA1_OBSERVA"))
            .OFEX_PRA1_ADVERTENCIA = NzVBA(rs("OFEX_PRA1_ADVERTENCIA"))
            ' Pruebas asistencia 2
            .PRA2_ID = NzLong(rs("PRA2_ID"))
            .OFEX_PRA2_PRE_COSTE = NzDouble(rs("OFEX_PRA2_PRE_COSTE"))
            .OFEX_PRA2_IMPORTE = NzDouble(rs("OFEX_PRA2_IMPORTE"))
            .OFEX_PRA2_DESCRIPCION = NzVBA(rs("OFEX_PRA2_DESCRIPCION"))
            .OFEX_PRA2_OBSERVA = NzVBA(rs("OFEX_PRA2_OBSERVA"))
            .OFEX_PRA2_ADVERTENCIA = NzVBA(rs("OFEX_PRA2_ADVERTENCIA"))
            ' Pruebas asistencia 3
            .PRA3_ID = NzLong(rs("PRA3_ID"))
            .OFEX_PRA3_PRE_COSTE = NzDouble(rs("OFEX_PRA3_PRE_COSTE"))
            .OFEX_PRA3_IMPORTE = NzDouble(rs("OFEX_PRA3_IMPORTE"))
            .OFEX_PRA3_DESCRIPCION = NzVBA(rs("OFEX_PRA3_DESCRIPCION"))
            .OFEX_PRA3_OBSERVA = NzVBA(rs("OFEX_PRA3_OBSERVA"))
            .OFEX_PRA3_ADVERTENCIA = NzVBA(rs("OFEX_PRA3_ADVERTENCIA"))
            ' Pruebas asistencia 4
            .PRA4_ID = NzLong(rs("PRA4_ID"))
            .OFEX_PRA4_PRE_COSTE = NzDouble(rs("OFEX_PRA4_PRE_COSTE"))
            .OFEX_PRA4_IMPORTE = NzDouble(rs("OFEX_PRA4_IMPORTE"))
            .OFEX_PRA4_DESCRIPCION = NzVBA(rs("OFEX_PRA4_DESCRIPCION"))
            .OFEX_PRA4_OBSERVA = NzVBA(rs("OFEX_PRA4_OBSERVA"))
            .OFEX_PRA4_ADVERTENCIA = NzVBA(rs("OFEX_PRA4_ADVERTENCIA"))

            ' Resolver FKs de extras
            If .EMBA_ID > 0 Then
                Dim emba As tCatalogEmbalajes
                emba = m_Catalogos.ObtenerEmbalajes(.EMBA_ID)
                If .OFEX_EMBA_DESCRIPCION = "" Then
                    .OFEX_EMBA_DESCRIPCION = emba.EMBA_MODELO & " - " & emba.EMBA_DESCRIPCION
                End If
                If .OFEX_EMBA_PRE_COSTE = 0 Then
                    .OFEX_EMBA_PRE_COSTE = emba.EMBA_PRE_COSTE
                End If
            End If

            If .TRAN_ID > 0 Then
                Dim tran As tCatalogTransportes
                tran = m_Catalogos.ObtenerTransportes(.TRAN_ID)
                If .OFEX_TRAN_DESCRIPCION = "" Then
                    .OFEX_TRAN_DESCRIPCION = tran.TRAN_MODELO & " - " & tran.TRAN_DESCRIPCION
                End If
                If .OFEX_TRAN_PRE_COSTE = 0 Then
                    .OFEX_TRAN_PRE_COSTE = tran.TRAN_PRE_COSTE
                End If
            End If

            If .PUMA_ID > 0 Then
                Dim puma As tCatalogPuestasMarcha
                puma = m_Catalogos.ObtenerPuestasMarcha(.PUMA_ID)
                If .OFEX_PUMA_DESCRIPCION = "" Then
                    .OFEX_PUMA_DESCRIPCION = puma.PUMA_MODELO & " - " & puma.PUMA_DESCRIPCION
                End If
                If .OFEX_PUMA_PRE_COSTE = 0 Then
                    .OFEX_PUMA_PRE_COSTE = puma.PUMA_PRE_COSTE
                End If
            End If

            If .PRA1_ID > 0 Then
                Dim pra1 As tCatalogPruebasAsistencia
                pra1 = m_Catalogos.ObtenerPruebasAsistencia(.PRA1_ID)
                If .OFEX_PRA1_DESCRIPCION = "" Then
                    .OFEX_PRA1_DESCRIPCION = pra1.PRAS_MODELO & " - " & pra1.PRAS_DESCRIPCION
                End If
                If .OFEX_PRA1_PRE_COSTE = 0 Then
                    .OFEX_PRA1_PRE_COSTE = pra1.PRAS_PRE_COSTE
                End If
            End If

            If .PRA2_ID > 0 Then
                Dim pra2 As tCatalogPruebasAsistencia
                pra2 = m_Catalogos.ObtenerPruebasAsistencia(.PRA2_ID)
                If .OFEX_PRA2_DESCRIPCION = "" Then
                    .OFEX_PRA2_DESCRIPCION = pra2.PRAS_MODELO & " - " & pra2.PRAS_DESCRIPCION
                End If
                If .OFEX_PRA2_PRE_COSTE = 0 Then
                    .OFEX_PRA2_PRE_COSTE = pra2.PRAS_PRE_COSTE
                End If
            End If

            If .PRA3_ID > 0 Then
                Dim pra3 As tCatalogPruebasAsistencia
                pra3 = m_Catalogos.ObtenerPruebasAsistencia(.PRA3_ID)
                If .OFEX_PRA3_DESCRIPCION = "" Then
                    .OFEX_PRA3_DESCRIPCION = pra3.PRAS_MODELO & " - " & pra3.PRAS_DESCRIPCION
                End If
                If .OFEX_PRA3_PRE_COSTE = 0 Then
                    .OFEX_PRA3_PRE_COSTE = pra3.PRAS_PRE_COSTE
                End If
            End If

            If .PRA4_ID > 0 Then
                Dim pra4 As tCatalogPruebasAsistencia
                pra4 = m_Catalogos.ObtenerPruebasAsistencia(.PRA4_ID)
                If .OFEX_PRA4_DESCRIPCION = "" Then
                    .OFEX_PRA4_DESCRIPCION = pra4.PRAS_MODELO & " - " & pra4.PRAS_DESCRIPCION
                End If
                If .OFEX_PRA4_PRE_COSTE = 0 Then
                    .OFEX_PRA4_PRE_COSTE = pra4.PRAS_PRE_COSTE
                End If
            End If

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

    Do While Not rs.EOF
        Dim it As New clsOfertaOtro
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
