VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCatalogos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================
'@Description: Gestor centralizado de catálogos con cache
'@Purpose: Evitar duplicación de datos de maestros en memoria
'@Pattern: Arrays+Dictionary para resolver limitación VBA UDTs
'@Updated: FASE 1 - 32 catálogos nuevos (Total: 47 catálogos)
'=========================================================
Option Explicit

'=========================================================
' ARRAYS Y DICTIONARIES PARA CATÁLOGOS (47 totales)
' Patrón: m_XXXArray() + m_XXXIndice (Dict Long->Long)
'=========================================================

' Catálogos originales (15)
Private m_chasisArray() As tCatalogChasis
Private m_chasisIndice As Object
Private m_chasisCount As Long

Private m_correasArray() As tCatalogCorreas
Private m_correasIndice As Object
Private m_correasCount As Long

Private m_cabezalesArray() As tCatalogCabezales
Private m_cabezalesIndice As Object
Private m_cabezalesCount As Long

Private m_cilindrosArray() As tCatalogCilindros
Private m_cilindrosIndice As Object
Private m_cilindrosCount As Long

Private m_motoresArray() As tCatalogMotores
Private m_motoresIndice As Object
Private m_motoresCount As Long
Private m_MotoresClassArray() As tCatalogMotoresClass
Private m_MotoresClassIndice As Object
Private m_MotoresClassCount As Long
Private m_MotoresEspecialesArray() As tCatalogMotoresEspeciales
Private m_MotoresEspecialesIndice As Object
Private m_MotoresEspecialesCount As Long
Private m_MotoresFabricaArray() As tCatalogMotoresFabrica
Private m_MotoresFabricaIndice As Object
Private m_MotoresFabricaCount As Long
Private m_MotoresNormativasArray() As tCatalogMotoresNormativas
Private m_MotoresNormativasIndice As Object
Private m_MotoresNormativasCount As Long
Private m_MotoresTensionArray() As tCatalogMotoresTension
Private m_MotoresTensionIndice As Object
Private m_MotoresTensionCount As Long

Private m_gasesArray() As tCatalogGases
Private m_gasesIndice As Object
Private m_gasesCount As Long

Private m_refrigeradoresArray() As tCatalogRefrigeradores
Private m_refrigeradoresIndice As Object
Private m_refrigeradoresCount As Long

Private m_calderinesArray() As tCatalogCalderines
Private m_calderinesIndice As Object
Private m_calderinesCount As Long

Private m_volumenesArray() As tCatalogVolumenes
Private m_volumenesIndice As Object
Private m_volumenesCount As Long

Private m_acoplamientosDirectoArray() As tCatalogAcoplamientosDirecto
Private m_acoplamientosDirectoIndice As Object
Private m_acoplamientosDirectoCount As Long

Private m_reductoresArray() As tCatalogReductores
Private m_reductoresIndice As Object
Private m_reductoresCount As Long

Private m_valvulasRetencionArray() As tCatalogValvulasRetencion
Private m_valvulasRetencionIndice As Object
Private m_valvulasRetencionCount As Long

Private m_electrovalvulasArray() As tCatalogElectrovalvulas
Private m_electrovalvulasIndice As Object
Private m_electrovalvulasCount As Long

Private m_filtrosArray() As tCatalogFiltros
Private m_filtrosIndice As Object
Private m_filtrosCount As Long

Private m_manometrosArray() As tCatalogManometros
Private m_manometrosIndice As Object
Private m_manometrosCount As Long

' Catálogos nuevos FASE 1 (32)
Private m_embalajesArray() As tCatalogEmbalajes
Private m_embalajesIndice As Object  ' Dictionary Long->Long
Private m_embalajesCount As Long
Private m_transportesArray() As tCatalogTransportes
Private m_transportesIndice As Object  ' Dictionary Long->Long
Private m_transportesCount As Long
Private m_puestasMarchaArray() As tCatalogPuestasMarcha
Private m_puestasMarchaIndice As Object  ' Dictionary Long->Long
Private m_puestasMarchaCount As Long
Private m_pruebasAsistenciaArray() As tCatalogPruebasAsistencia
Private m_pruebasAsistenciaIndice As Object  ' Dictionary Long->Long
Private m_pruebasAsistenciaCount As Long
Private m_aerosArray() As tCatalogAeros
Private m_aerosIndice As Object  ' Dictionary Long->Long
Private m_aerosCount As Long
Private m_cajaLocalArray() As tCatalogCajaLocal
Private m_cajaLocalIndice As Object  ' Dictionary Long->Long
Private m_cajaLocalCount As Long
Private m_gruposEngraseArray() As tCatalogGruposEngrase
Private m_gruposEngraseIndice As Object  ' Dictionary Long->Long
Private m_gruposEngraseCount As Long
Private m_llavesEntradaArray() As tCatalogLlavesEntrada
Private m_llavesEntradaIndice As Object  ' Dictionary Long->Long
Private m_llavesEntradaCount As Long
Private m_llavesSalidaArray() As tCatalogLlavesSalida
Private m_llavesSalidaIndice As Object  ' Dictionary Long->Long
Private m_llavesSalidaCount As Long
Private m_normativasCompresorArray() As tCatalogNormativasCompresor
Private m_normativasCompresorIndice As Object  ' Dictionary Long->Long
Private m_normativasCompresorCount As Long
Private m_valvulasPresionArray() As tCatalogValvulasPresion
Private m_valvulasPresionIndice As Object  ' Dictionary Long->Long
Private m_valvulasPresionCount As Long
Private m_engraseCilindrosArray() As tCatalogEngraseCilindros
Private m_engraseCilindrosIndice As Object  ' Dictionary Long->Long
Private m_engraseCilindrosCount As Long
Private m_purgadoresArray() As tCatalogPurgadores
Private m_purgadoresIndice As Object  ' Dictionary Long->Long
Private m_purgadoresCount As Long
Private m_resistenciasCalefaccionArray() As tCatalogResistenciasCalefaccion
Private m_resistenciasCalefaccionIndice As Object  ' Dictionary Long->Long
Private m_resistenciasCalefaccionCount As Long
Private m_valvulasReguladorasArray() As tCatalogValvulasReguladoras
Private m_valvulasReguladorasIndice As Object  ' Dictionary Long->Long
Private m_valvulasReguladorasCount As Long
Private m_paisesArray() As tCatalogPaises
Private m_paisesIndice As Object  ' Dictionary Long->Long
Private m_paisesCount As Long
Private m_comercialesArray() As tCatalogComerciales
Private m_comercialesIndice As Object  ' Dictionary Long->Long
Private m_comercialesCount As Long
Private m_ofertantesArray() As tCatalogOfertantes
Private m_ofertantesIndice As Object  ' Dictionary Long->Long
Private m_ofertantesCount As Long
Private m_arrancadoresFuerzaArray() As tCatalogArrancadoresFuerza
Private m_arrancadoresFuerzaIndice As Object  ' Dictionary Long->Long
Private m_arrancadoresFuerzaCount As Long
Private m_arrancadoresFuerzaControlArray() As tCatalogArrancadoresFuerzaControl
Private m_arrancadoresFuerzaControlIndice As Object  ' Dictionary Long->Long
Private m_arrancadoresFuerzaControlCount As Long
Private m_arrancadoresControlArray() As tCatalogArrancadoresControl
Private m_arrancadoresControlIndice As Object  ' Dictionary Long->Long
Private m_arrancadoresControlCount As Long
Private m_arrancadoresFuerzaControlPETArray() As tCatalogArrancadoresFuerzaControlPET
Private m_arrancadoresFuerzaControlPETIndice As Object  ' Dictionary Long->Long
Private m_arrancadoresFuerzaControlPETCount As Long
Private m_transmisoresTemperaturaArray() As tCatalogTransmisoresTemperatura
Private m_transmisoresTemperaturaIndice As Object  ' Dictionary Long->Long
Private m_transmisoresTemperaturaCount As Long
Private m_transmisoresPresionArray() As tCatalogTransmisoresPresion
Private m_transmisoresPresionIndice As Object  ' Dictionary Long->Long
Private m_transmisoresPresionCount As Long
Private m_termometrosArray() As tCatalogTermometros
Private m_termometrosIndice As Object  ' Dictionary Long->Long
Private m_termometrosCount As Long
Private m_electrovalvulasRegulacionArray() As tCatalogElectrovalvulasRegulacion
Private m_electrovalvulasRegulacionIndice As Object  ' Dictionary Long->Long
Private m_electrovalvulasRegulacionCount As Long
Private m_sensoresCaidaVastagoArray() As tCatalogSensoresCaidaVastago
Private m_sensoresCaidaVastagoIndice As Object  ' Dictionary Long->Long
Private m_sensoresCaidaVastagoCount As Long
Private m_interruptoresVibracionArray() As tCatalogInterruptoresVibracion
Private m_interruptoresVibracionIndice As Object  ' Dictionary Long->Long
Private m_interruptoresVibracionCount As Long
Private m_interruptoresNivelAceiteArray() As tCatalogInterruptoresNivelAceite
Private m_interruptoresNivelAceiteIndice As Object  ' Dictionary Long->Long
Private m_interruptoresNivelAceiteCount As Long
Private m_nivelesCondensadosArray() As tCatalogNivelesCondensados
Private m_nivelesCondensadosIndice As Object  ' Dictionary Long->Long
Private m_nivelesCondensadosCount As Long
Private m_valvulasTermostaticasArray() As tCatalogValvulasTermostaticas
Private m_valvulasTermostaticasIndice As Object  ' Dictionary Long->Long
Private m_valvulasTermostaticasCount As Long
Private m_instrumentacionesArray() As tCatalogInstrumentaciones
Private m_instrumentacionesIndice As Object  ' Dictionary Long->Long
Private m_instrumentacionesCount As Long

Private m_tiposMaterialArray() As tCatalogTiposMaterial
Private m_tiposMaterialIndice As Object
Private m_tiposMaterialCount As Long

Private m_calderinesPredeterminadosArray() As tCatalogCalderinesPredeterminados
Private m_calderinesPredeterminadosIndice As Object
Private m_calderinesPredeterminadosCount As Long

Private m_manoObraCabezalesArray() As tCatalogManoObraCabezales
Private m_manoObraCabezalesIndice As Object
Private m_manoObraCabezalesCount As Long

Private m_manoObraIngenieriasArray() As tCatalogManoObraIngenierias
Private m_manoObraIngenieriasIndice As Object
Private m_manoObraIngenieriasCount As Long

Private m_CalderinesCostesArray() As tCatalogCalderinesCostes
Private m_CalderinesCostesIndice As Object
Private m_CalderinesCostesCount As Long
Private m_DepositosEntradaArray() As tCatalogDepositosEntrada
Private m_DepositosEntradaIndice As Object
Private m_DepositosEntradaCount As Long
Private m_DepositosFinalArray() As tCatalogDepositosFinal
Private m_DepositosFinalIndice As Object
Private m_DepositosFinalCount As Long

Private m_AreasArray() As tCatalogAreas
Private m_AreasIndice As Object
Private m_AreasCount As Long

Private m_ParametrosArray() As tCatalogParametros
Private m_ParametrosIndice As Object
Private m_ParametrosCount As Long


Private m_RefrigeradoresCostesArray() As tCatalogRefrigeradoresCostes
Private m_RefrigeradoresCostesIndice As Object
Private m_RefrigeradoresCostesCount As Long
Private m_RefrigeradoresEspecialesArray() As tCatalogRefrigeradoresEspeciales
Private m_RefrigeradoresEspecialesIndice As Object
Private m_RefrigeradoresEspecialesCount As Long
Private m_TiposPresionNominalArray() As tCatalogTiposPresionNominal
Private m_TiposPresionNominalIndice As Object
Private m_TiposPresionNominalCount As Long
Private m_TiposTubosArray() As tCatalogTiposTubos
Private m_TiposTubosIndice As Object
Private m_TiposTubosCount As Long
Private m_TuberiasAguaArray() As tCatalogTuberiasAgua
Private m_TuberiasAguaIndice As Object
Private m_TuberiasAguaCount As Long
Private m_TuberiasAireArray() As tCatalogTuberiasAire
Private m_TuberiasAireIndice As Object
Private m_TuberiasAireCount As Long
Private m_TuberiasAireDiametrosArray() As tCatalogTuberiasAireDiametros
Private m_TuberiasAireDiametrosIndice As Object
Private m_TuberiasAireDiametrosCount As Long
Private m_ValvulasSeguridadArray() As tCatalogValvulasSeguridad
Private m_ValvulasSeguridadIndice As Object
Private m_ValvulasSeguridadCount As Long




Private m_Ctx As clsDBContext
Private m_Cargado As Boolean

'=========================================================
' INICIALIZACIÓN Y CARGA
'=========================================================

Private Sub Class_Initialize()
    m_Cargado = False
End Sub

'@Description: Carga todos los catálogos en memoria desde la BD
'@Param ctx: Contexto de base de datos ya conectado
Public Sub CargarTodos(ByVal ctx As clsDBContext)
    On Error GoTo ErrHandler
    
    If m_Cargado Then Exit Sub
    Set m_Ctx = ctx
    
    ' Cargar catálogos originales (15)
    CargarChasis
    CargarCorreas
    CargarCabezales
    CargarCilindros
    CargarMotores
    
    CargarMotoresClass
    CargarMotoresEspeciales
    CargarMotoresFabrica
    CargarMotoresNormativas
    CargarMotoresTension
    
    CargarGases
    CargarRefrigeradores
    CargarRefrigeradoresCostes
    CargarRefrigeradoresEspeciales
    CargarTiposPresionNominal
    CargarTiposTubos
    CargarTuberiasAgua
    CargarTuberiasAire
    CargarTuberiasAireDiametros
    CargarValvulasSeguridad
    CargarCalderines
    CargarCalderinesCostes
    CargarDepositosEntrada
    CargarDepositosFinal
    CargarVolumenes
    CargarAcoplamientosDirecto
    CargarReductores
    CargarValvulasRetencion
    CargarElectrovalvulas
    CargarFiltros
    CargarManometros
    
    ' Cargar catálogos FASE 1 (32)
    CargarEmbalajes
    CargarTransportes
    CargarPuestasMarcha
    CargarPruebasAsistencia
    CargarAeros
    CargarCajaLocal
    CargarGruposEngrase
    CargarLlavesEntrada
    CargarLlavesSalida
    CargarNormativasCompresor
    CargarValvulasPresion
    CargarEngraseCilindros
    CargarPurgadores
    CargarResistenciasCalefaccion
    CargarValvulasReguladoras
    CargarPaises
    CargarComerciales
    CargarOfertantes
    CargarArrancadoresFuerza
    CargarArrancadoresFuerzaControl
    CargarArrancadoresControl
    CargarArrancadoresFuerzaControlPET
    CargarTransmisoresTemperatura
    CargarTransmisoresPresion
    CargarTermometros
    CargarElectrovalvulasRegulacion
    CargarSensoresCaidaVastago
    CargarInterruptoresVibracion
    CargarInterruptoresNivelAceite
    CargarNivelesCondensados
    CargarValvulasTermostaticas
    CargarInstrumentaciones
    
   CargarTiposMaterial
   CargarCalderinesPredeterminados
   CargarManoObraCabezales
   CargarManoObraIngenierias
    
    m_Cargado = True
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "clsCatalogos.CargarTodos", _
        "Error cargando catálogos: " & Err.Description
End Sub

'=========================================================
' MÉTODOS PRIVADOS DE CARGA - CATÁLOGOS ORIGINALES (15)
'=========================================================

Public Sub CargarChasis()
    On Error GoTo ErrHandler
    
    Set m_chasisIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Chasis")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_chasisCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_chasisArray(1 To m_chasisCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_chasisArray(i).CHAS_ID = NzLong(rs("CHAS_ID"))
        m_chasisArray(i).CHAS_MODELO = NzVBA(rs("CHAS_MODELO"))
        m_chasisArray(i).CHAS_DESCRIPCION = NzVBA(rs("CHAS_DESCRIPCION"))
        m_chasisArray(i).CHAS_ADVERTENCIA = NzVBA(rs("CHAS_ADVERTENCIA"))
        m_chasisArray(i).CHAS_PRE_COSTE = NzDouble(rs("CHAS_PRE_COSTE"))
        
        m_chasisIndice.Add m_chasisArray(i).CHAS_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarChasis", "Error: " & Err.Description
End Sub

Public Sub CargarCorreas()
    On Error GoTo ErrHandler
    
    Set m_correasIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Correas")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_correasCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_correasArray(1 To m_correasCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_correasArray(i).CORR_ID = NzLong(rs("CORR_ID"))
        m_correasArray(i).CORR_MODELO = NzVBA(rs("CORR_MODELO"))
        m_correasArray(i).CORR_DESCRIPCION = NzVBA(rs("CORR_DESCRIPCION"))
        m_correasArray(i).CORR_ADVERTENCIA = NzVBA(rs("CORR_ADVERTENCIA"))
        m_correasArray(i).CORR_PRE_COSTE = NzDouble(rs("CORR_PRE_COSTE"))
        m_correasArray(i).CORR_ATEX_COSTE = NzDouble(rs("CORR_ATEX_COSTE"))
        
        m_correasIndice.Add m_correasArray(i).CORR_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarCorreas", "Error: " & Err.Description
End Sub

Public Sub CargarCabezales()
    On Error GoTo ErrHandler
    
    Set m_cabezalesIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Cabezales")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_cabezalesCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_cabezalesArray(1 To m_cabezalesCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_cabezalesArray(i).CABE_ID = NzLong(rs("CABE_ID"))
        m_cabezalesArray(i).CABE_MODELO = NzVBA(rs("CABE_MODELO"))
        m_cabezalesArray(i).CABE_DESCRIPCION = NzVBA(rs("CABE_DESCRIPCION"))
        m_cabezalesArray(i).CABE_ADVERTENCIA = NzVBA(rs("CABE_ADVERTENCIA"))
        m_cabezalesArray(i).CABE_PRE_COSTE = NzDouble(rs("CABE_PRE_COSTE"))
        m_cabezalesArray(i).CABE_BLOQUE_SAS = NzVBA(rs("CABE_BLOQUE_SAS"))
        
        m_cabezalesIndice.Add m_cabezalesArray(i).CABE_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarCabezales", "Error: " & Err.Description
End Sub

Public Sub CargarCilindros()
    On Error GoTo ErrHandler
    
    Set m_cilindrosIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Cilindros")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_cilindrosCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_cilindrosArray(1 To m_cilindrosCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_cilindrosArray(i).CILI_ID = NzLong(rs("CILI_ID"))
        m_cilindrosArray(i).CILI_MODELO = NzVBA(rs("CILI_MODELO"))
        m_cilindrosArray(i).CILI_DESCRIPCION = NzVBA(rs("CILI_DESCRIPCION"))
        m_cilindrosArray(i).CILI_ADVERTENCIA = NzVBA(rs("CILI_ADVERTENCIA"))
        m_cilindrosArray(i).CILI_PRE_COSTE = NzDouble(rs("CILI_PRE_COSTE"))
        m_cilindrosArray(i).CILI_T2_COSTE = NzDouble(rs("CILI_T2_COSTE"))
        m_cilindrosArray(i).CILI_CO2_COSTE = NzDouble(rs("CILI_CO2_COSTE"))
        m_cilindrosArray(i).CILI_CPI_COSTE = NzDouble(rs("CILI_CPI_COSTE"))
        m_cilindrosArray(i).CILI_SAS_COSTE = NzDouble(rs("CILI_SAS_COSTE"))
        m_cilindrosArray(i).CILI_EMP_COSTE = NzDouble(rs("CILI_EMP_COSTE"))
        
        m_cilindrosIndice.Add m_cilindrosArray(i).CILI_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarCilindros", "Error: " & Err.Description
End Sub

Public Sub CargarMotores()
    On Error GoTo ErrHandler
    
    Set m_motoresIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Motores")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_motoresCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_motoresArray(1 To m_motoresCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_motoresArray(i).MOTO_ID = NzLong(rs("MOTO_ID"))
        m_motoresArray(i).MOTO_MODELO = NzVBA(rs("MOTO_MODELO"))
        m_motoresArray(i).MOTO_DESCRIPCION = NzVBA(rs("MOTO_DESCRIPCION"))
        m_motoresArray(i).MOTO_ADVERTENCIA = NzVBA(rs("MOTO_ADVERTENCIA"))
        m_motoresArray(i).MOTO_PRE_COSTE = NzDouble(rs("MOTO_PRE_COSTE"))
        m_motoresArray(i).MTNO_ID = NzLong(rs("MTNO_ID"))
        m_motoresArray(i).MTCL_ID = NzLong(rs("MTCL_ID"))
        m_motoresArray(i).MTFB_ID = NzLong(rs("MTFB_ID"))
        m_motoresArray(i).MTTN_ID = NzLong(rs("MTTN_ID"))
        
        m_motoresIndice.Add m_motoresArray(i).MOTO_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarMotores", "Error: " & Err.Description
End Sub

Public Sub CargarGases()
    On Error GoTo ErrHandler
    
    Set m_gasesIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Gases")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_gasesCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_gasesArray(1 To m_gasesCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_gasesArray(i).GASE_ID = NzLong(rs("GASE_ID"))
        m_gasesArray(i).GASE_DENOMINACION = NzVBA(rs("GASE_DENOMINACION"))
        
        m_gasesIndice.Add m_gasesArray(i).GASE_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarGases", "Error: " & Err.Description
End Sub

Public Sub CargarRefrigeradores()
    On Error GoTo ErrHandler
    
    Set m_refrigeradoresIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Refrigeradores")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_refrigeradoresCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_refrigeradoresArray(1 To m_refrigeradoresCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_refrigeradoresArray(i).REFR_ID = NzLong(rs("REFR_ID"))
        m_refrigeradoresArray(i).REFR_MODELO = NzVBA(rs("REFR_MODELO"))
        m_refrigeradoresArray(i).REFR_DESCRIPCION = NzVBA(rs("REFR_DESCRIPCION"))
        m_refrigeradoresArray(i).REFR_ADVERTENCIA = NzVBA(rs("REFR_ADVERTENCIA"))
        
        m_refrigeradoresIndice.Add m_refrigeradoresArray(i).REFR_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarRefrigeradores", "Error: " & Err.Description
End Sub

Public Sub CargarCalderines()
    On Error GoTo ErrHandler
    
    Set m_calderinesIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Calderines")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_calderinesCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_calderinesArray(1 To m_calderinesCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_calderinesArray(i).CALD_ID = NzLong(rs("CALD_ID"))
        m_calderinesArray(i).CALD_MODELO = NzVBA(rs("CALD_MODELO"))
        m_calderinesArray(i).CALD_DESCRIPCION = NzVBA(rs("CALD_DESCRIPCION"))
        m_calderinesArray(i).CALD_ADVERTENCIA = NzVBA(rs("CALD_ADVERTENCIA"))
        
        m_calderinesIndice.Add m_calderinesArray(i).CALD_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarCalderines", "Error: " & Err.Description
End Sub

Public Sub CargarVolumenes()
    On Error GoTo ErrHandler
    
    Set m_volumenesIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Volumenes")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_volumenesCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_volumenesArray(1 To m_volumenesCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_volumenesArray(i).VOLU_ID = NzLong(rs("VOLU_ID"))
        m_volumenesArray(i).VOLU_CANTIDAD = NzVBA(rs("VOLU_CANTIDAD"))
        
        m_volumenesIndice.Add m_volumenesArray(i).VOLU_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarVolumenes", "Error: " & Err.Description
End Sub

Public Sub CargarAcoplamientosDirecto()
    On Error GoTo ErrHandler
    
    Set m_acoplamientosDirectoIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM AcoplamientosDirecto")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_acoplamientosDirectoCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_acoplamientosDirectoArray(1 To m_acoplamientosDirectoCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_acoplamientosDirectoArray(i).ACOD_ID = NzLong(rs("ACOD_ID"))
        m_acoplamientosDirectoArray(i).ACOD_MODELO = NzVBA(rs("ACOD_MODELO"))
        m_acoplamientosDirectoArray(i).ACOD_DESCRIPCION = NzVBA(rs("ACOD_DESCRIPCION"))
        m_acoplamientosDirectoArray(i).ACOD_ADVERTENCIA = NzVBA(rs("ACOD_ADVERTENCIA"))
        m_acoplamientosDirectoArray(i).ACOD_PRE_COSTE = NzDouble(rs("ACOD_PRE_COSTE"))
        
        m_acoplamientosDirectoIndice.Add m_acoplamientosDirectoArray(i).ACOD_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarAcoplamientosDirecto", "Error: " & Err.Description
End Sub

Public Sub CargarReductores()
    On Error GoTo ErrHandler
    
    Set m_reductoresIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Reductores")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_reductoresCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_reductoresArray(1 To m_reductoresCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_reductoresArray(i).REDU_ID = NzLong(rs("REDU_ID"))
        m_reductoresArray(i).REDU_MODELO = NzVBA(rs("REDU_MODELO"))
        m_reductoresArray(i).REDU_DESCRIPCION = NzVBA(rs("REDU_DESCRIPCION"))
        m_reductoresArray(i).REDU_ADVERTENCIA = NzVBA(rs("REDU_ADVERTENCIA"))
        m_reductoresArray(i).REDU_PRE_COSTE = NzDouble(rs("REDU_PRE_COSTE"))
        m_reductoresArray(i).REDU_ATEX_COSTE = NzDouble(rs("REDU_ATEX_COSTE"))
        
        m_reductoresIndice.Add m_reductoresArray(i).REDU_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarReductores", "Error: " & Err.Description
End Sub

Public Sub CargarValvulasRetencion()
    On Error GoTo ErrHandler
    
    Set m_valvulasRetencionIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM ValvulasRetencion")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_valvulasRetencionCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_valvulasRetencionArray(1 To m_valvulasRetencionCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_valvulasRetencionArray(i).VARE_ID = NzLong(rs("VARE_ID"))
        m_valvulasRetencionArray(i).VARE_MODELO = NzVBA(rs("VARE_MODELO"))
        m_valvulasRetencionArray(i).VARE_DESCRIPCION = NzVBA(rs("VARE_DESCRIPCION"))
        m_valvulasRetencionArray(i).VARE_ADVERTENCIA = NzVBA(rs("VARE_ADVERTENCIA"))
        m_valvulasRetencionArray(i).VARE_PRE_COSTE = NzDouble(rs("VARE_PRE_COSTE"))
        
        m_valvulasRetencionIndice.Add m_valvulasRetencionArray(i).VARE_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarValvulasRetencion", "Error: " & Err.Description
End Sub

Public Sub CargarElectrovalvulas()
    On Error GoTo ErrHandler
    
    Set m_electrovalvulasIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Electrovalvulas")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_electrovalvulasCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_electrovalvulasArray(1 To m_electrovalvulasCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_electrovalvulasArray(i).ELEC_ID = NzLong(rs("ELEC_ID"))
        m_electrovalvulasArray(i).ELEC_MODELO = NzVBA(rs("ELEC_MODELO"))
        m_electrovalvulasArray(i).ELEC_DESCRIPCION = NzVBA(rs("ELEC_DESCRIPCION"))
        m_electrovalvulasArray(i).ELEC_ADVERTENCIA = NzVBA(rs("ELEC_ADVERTENCIA"))
        m_electrovalvulasArray(i).ELEC_PRE_COSTE = NzDouble(rs("ELEC_PRE_COSTE"))
        
        m_electrovalvulasIndice.Add m_electrovalvulasArray(i).ELEC_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarElectrovalvulas", "Error: " & Err.Description
End Sub

Public Sub CargarFiltros()
    On Error GoTo ErrHandler
    
    Set m_filtrosIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Filtros")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_filtrosCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_filtrosArray(1 To m_filtrosCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_filtrosArray(i).FILT_ID = NzLong(rs("FILT_ID"))
        m_filtrosArray(i).FILT_MODELO = NzVBA(rs("FILT_MODELO"))
        m_filtrosArray(i).FILT_DESCRIPCION = NzVBA(rs("FILT_DESCRIPCION"))
        m_filtrosArray(i).FILT_ADVERTENCIA = NzVBA(rs("FILT_ADVERTENCIA"))
        m_filtrosArray(i).FILT_PRE_COSTE = NzDouble(rs("FILT_PRE_COSTE"))
        
        m_filtrosIndice.Add m_filtrosArray(i).FILT_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarFiltros", "Error: " & Err.Description
End Sub

Public Sub CargarManometros()
    On Error GoTo ErrHandler
    
    Set m_manometrosIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Manometros")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_manometrosCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_manometrosArray(1 To m_manometrosCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_manometrosArray(i).MANO_ID = NzLong(rs("MANO_ID"))
        m_manometrosArray(i).MANO_MODELO = NzVBA(rs("MANO_MODELO"))
        m_manometrosArray(i).MANO_DESCRIPCION = NzVBA(rs("MANO_DESCRIPCION"))
        m_manometrosArray(i).MANO_ADVERTENCIA = NzVBA(rs("MANO_ADVERTENCIA"))
        m_manometrosArray(i).MANO_PRE_COSTE = NzDouble(rs("MANO_PRE_COSTE"))
        
        m_manometrosIndice.Add m_manometrosArray(i).MANO_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarManometros", "Error: " & Err.Description
End Sub


'=========================================================
' MÉTODOS PRIVADOS DE CARGA - CATÁLOGOS FASE 1 (32)
'=========================================================

Public Sub CargarEmbalajes()
    On Error GoTo ErrHandler
    
    Set m_embalajesIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Embalajes")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_embalajesCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_embalajesArray(1 To m_embalajesCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_embalajesArray(i).EMBA_ID = NzLong(rs("EMBA_ID"))
        m_embalajesArray(i).EMBA_MODELO = NzVBA(rs("EMBA_MODELO"))
        m_embalajesArray(i).EMBA_DESCRIPCION = NzVBA(rs("EMBA_DESCRIPCION"))
        m_embalajesArray(i).EMBA_ADVERTENCIA = NzVBA(rs("EMBA_ADVERTENCIA"))
        m_embalajesArray(i).EMBA_PRE_COSTE = NzDouble(rs("EMBA_PRE_COSTE"))
        
        ' Agregar al índice
        m_embalajesIndice.Add m_embalajesArray(i).EMBA_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarEmbalajes", _
        "Error: " & Err.Description
End Sub

Public Sub CargarTransportes()
    On Error GoTo ErrHandler
    
    Set m_transportesIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Transportes")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_transportesCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_transportesArray(1 To m_transportesCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_transportesArray(i).TRAN_ID = NzLong(rs("TRAN_ID"))
        m_transportesArray(i).TRAN_MODELO = NzVBA(rs("TRAN_MODELO"))
        m_transportesArray(i).TRAN_DESCRIPCION = NzVBA(rs("TRAN_DESCRIPCION"))
        m_transportesArray(i).TRAN_ADVERTENCIA = NzVBA(rs("TRAN_ADVERTENCIA"))
        m_transportesArray(i).TRAN_PRE_COSTE = NzDouble(rs("TRAN_PRE_COSTE"))
        
        ' Agregar al índice
        m_transportesIndice.Add m_transportesArray(i).TRAN_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarTransportes", _
        "Error: " & Err.Description
End Sub

Public Sub CargarPuestasMarcha()
    On Error GoTo ErrHandler
    
    Set m_puestasMarchaIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM PuestasMarcha")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_puestasMarchaCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_puestasMarchaArray(1 To m_puestasMarchaCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_puestasMarchaArray(i).PUMA_ID = NzLong(rs("PUMA_ID"))
        m_puestasMarchaArray(i).PUMA_MODELO = NzVBA(rs("PUMA_MODELO"))
        m_puestasMarchaArray(i).PUMA_DESCRIPCION = NzVBA(rs("PUMA_DESCRIPCION"))
        m_puestasMarchaArray(i).PUMA_ADVERTENCIA = NzVBA(rs("PUMA_ADVERTENCIA"))
        m_puestasMarchaArray(i).PUMA_PRE_COSTE = NzDouble(rs("PUMA_PRE_COSTE"))
        
        ' Agregar al índice
        m_puestasMarchaIndice.Add m_puestasMarchaArray(i).PUMA_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarPuestasMarcha", _
        "Error: " & Err.Description
End Sub

Public Sub CargarPruebasAsistencia()
    On Error GoTo ErrHandler
    
    Set m_pruebasAsistenciaIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM PruebasAsistencia")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_pruebasAsistenciaCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_pruebasAsistenciaArray(1 To m_pruebasAsistenciaCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_pruebasAsistenciaArray(i).PRAS_ID = NzLong(rs("PRAS_ID"))
        m_pruebasAsistenciaArray(i).PRAS_MODELO = NzVBA(rs("PRAS_MODELO"))
        m_pruebasAsistenciaArray(i).PRAS_DESCRIPCION = NzVBA(rs("PRAS_DESCRIPCION"))
        m_pruebasAsistenciaArray(i).PRAS_ADVERTENCIA = NzVBA(rs("PRAS_ADVERTENCIA"))
        m_pruebasAsistenciaArray(i).PRAS_PRE_COSTE = NzDouble(rs("PRAS_PRE_COSTE"))
        
        ' Agregar al índice
        m_pruebasAsistenciaIndice.Add m_pruebasAsistenciaArray(i).PRAS_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarPruebasAsistencia", _
        "Error: " & Err.Description
End Sub

Public Sub CargarAeros()
    On Error GoTo ErrHandler
    
    Set m_aerosIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Aeros")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_aerosCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_aerosArray(1 To m_aerosCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_aerosArray(i).AERO_ID = NzLong(rs("AERO_ID"))
        m_aerosArray(i).AERO_MODELO = NzVBA(rs("AERO_MODELO"))
        m_aerosArray(i).AERO_DESCRIPCION = NzVBA(rs("AERO_DESCRIPCION"))
        m_aerosArray(i).AERO_ADVERTENCIA = NzVBA(rs("AERO_ADVERTENCIA"))
        m_aerosArray(i).AERO_PRE_COSTE = NzDouble(rs("AERO_PRE_COSTE"))
        
        ' Agregar al índice
        m_aerosIndice.Add m_aerosArray(i).AERO_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarAeros", _
        "Error: " & Err.Description
End Sub

Public Sub CargarCajaLocal()
    On Error GoTo ErrHandler
    
    Set m_cajaLocalIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM CajaLocal")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_cajaLocalCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_cajaLocalArray(1 To m_cajaLocalCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_cajaLocalArray(i).CALO_ID = NzLong(rs("CALO_ID"))
        m_cajaLocalArray(i).CALO_MODELO = NzVBA(rs("CALO_MODELO"))
        m_cajaLocalArray(i).CALO_DESCRIPCION = NzVBA(rs("CALO_DESCRIPCION"))
        m_cajaLocalArray(i).CALO_ADVERTENCIA = NzVBA(rs("CALO_ADVERTENCIA"))
        m_cajaLocalArray(i).CALO_PRE_COSTE = NzDouble(rs("CALO_PRE_COSTE"))
        
        ' Agregar al índice
        m_cajaLocalIndice.Add m_cajaLocalArray(i).CALO_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarCajaLocal", _
        "Error: " & Err.Description
End Sub

Public Sub CargarGruposEngrase()
    On Error GoTo ErrHandler
    
    Set m_gruposEngraseIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM GruposEngrase")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_gruposEngraseCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_gruposEngraseArray(1 To m_gruposEngraseCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_gruposEngraseArray(i).GREN_ID = NzLong(rs("GREN_ID"))
        m_gruposEngraseArray(i).GREN_MODELO = NzVBA(rs("GREN_MODELO"))
        m_gruposEngraseArray(i).GREN_DESCRIPCION = NzVBA(rs("GREN_DESCRIPCION"))
        m_gruposEngraseArray(i).GREN_ADVERTENCIA = NzVBA(rs("GREN_ADVERTENCIA"))
        m_gruposEngraseArray(i).GREN_PRE_COSTE = NzDouble(rs("GREN_PRE_COSTE"))
        
        ' Agregar al índice
        m_gruposEngraseIndice.Add m_gruposEngraseArray(i).GREN_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarGruposEngrase", _
        "Error: " & Err.Description
End Sub

Public Sub CargarLlavesEntrada()
    On Error GoTo ErrHandler
    
    Set m_llavesEntradaIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM LlavesEntrada")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_llavesEntradaCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_llavesEntradaArray(1 To m_llavesEntradaCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_llavesEntradaArray(i).LLEN_ID = NzLong(rs("LLEN_ID"))
        m_llavesEntradaArray(i).LLEN_MODELO = NzVBA(rs("LLEN_MODELO"))
        m_llavesEntradaArray(i).LLEN_DESCRIPCION = NzVBA(rs("LLEN_DESCRIPCION"))
        m_llavesEntradaArray(i).LLEN_ADVERTENCIA = NzVBA(rs("LLEN_ADVERTENCIA"))
        m_llavesEntradaArray(i).LLEN_PRE_COSTE = NzDouble(rs("LLEN_PRE_COSTE"))
        
        ' Agregar al índice
        m_llavesEntradaIndice.Add m_llavesEntradaArray(i).LLEN_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarLlavesEntrada", _
        "Error: " & Err.Description
End Sub

Public Sub CargarLlavesSalida()
    On Error GoTo ErrHandler
    
    Set m_llavesSalidaIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM LlavesSalida")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_llavesSalidaCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_llavesSalidaArray(1 To m_llavesSalidaCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_llavesSalidaArray(i).LLSA_ID = NzLong(rs("LLSA_ID"))
        m_llavesSalidaArray(i).LLSA_MODELO = NzVBA(rs("LLSA_MODELO"))
        m_llavesSalidaArray(i).LLSA_DESCRIPCION = NzVBA(rs("LLSA_DESCRIPCION"))
        m_llavesSalidaArray(i).LLSA_ADVERTENCIA = NzVBA(rs("LLSA_ADVERTENCIA"))
        m_llavesSalidaArray(i).LLSA_PRE_COSTE = NzDouble(rs("LLSA_PRE_COSTE"))
        
        ' Agregar al índice
        m_llavesSalidaIndice.Add m_llavesSalidaArray(i).LLSA_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarLlavesSalida", _
        "Error: " & Err.Description
End Sub

Public Sub CargarNormativasCompresor()
    On Error GoTo ErrHandler
    
    Set m_normativasCompresorIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM NormativasCompresor")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_normativasCompresorCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_normativasCompresorArray(1 To m_normativasCompresorCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_normativasCompresorArray(i).NORC_ID = NzLong(rs("NORC_ID"))
        m_normativasCompresorArray(i).NORC_MODELO = NzVBA(rs("NORC_MODELO"))
        m_normativasCompresorArray(i).NORC_DESCRIPCION = NzVBA(rs("NORC_DESCRIPCION"))
        m_normativasCompresorArray(i).NORC_ADVERTENCIA = NzVBA(rs("NORC_ADVERTENCIA"))
        m_normativasCompresorArray(i).NORC_PRE_COSTE = NzDouble(rs("NORC_PRE_COSTE"))
        
        ' Agregar al índice
        m_normativasCompresorIndice.Add m_normativasCompresorArray(i).NORC_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarNormativasCompresor", _
        "Error: " & Err.Description
End Sub

Public Sub CargarValvulasPresion()
    On Error GoTo ErrHandler
    
    Set m_valvulasPresionIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM ValvulasPresion")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_valvulasPresionCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_valvulasPresionArray(1 To m_valvulasPresionCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_valvulasPresionArray(i).VAPR_ID = NzLong(rs("VAPR_ID"))
        m_valvulasPresionArray(i).VAPR_MODELO = NzVBA(rs("VAPR_MODELO"))
        m_valvulasPresionArray(i).VAPR_DESCRIPCION = NzVBA(rs("VAPR_DESCRIPCION"))
        m_valvulasPresionArray(i).VAPR_ADVERTENCIA = NzVBA(rs("VAPR_ADVERTENCIA"))
        m_valvulasPresionArray(i).VAPR_PRE_COSTE = NzDouble(rs("VAPR_PRE_COSTE"))
        
        ' Agregar al índice
        m_valvulasPresionIndice.Add m_valvulasPresionArray(i).VAPR_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarValvulasPresion", _
        "Error: " & Err.Description
End Sub

Public Sub CargarEngraseCilindros()
    On Error GoTo ErrHandler
    
    Set m_engraseCilindrosIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM EngraseCilindros")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_engraseCilindrosCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_engraseCilindrosArray(1 To m_engraseCilindrosCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_engraseCilindrosArray(i).ENCI_ID = NzLong(rs("ENCI_ID"))
        m_engraseCilindrosArray(i).ENCI_MODELO = NzVBA(rs("ENCI_MODELO"))
        m_engraseCilindrosArray(i).ENCI_DESCRIPCION = NzVBA(rs("ENCI_DESCRIPCION"))
        m_engraseCilindrosArray(i).ENCI_ADVERTENCIA = NzVBA(rs("ENCI_ADVERTENCIA"))
        m_engraseCilindrosArray(i).ENCI_PRE_COSTE = NzDouble(rs("ENCI_PRE_COSTE"))
        
        ' Agregar al índice
        m_engraseCilindrosIndice.Add m_engraseCilindrosArray(i).ENCI_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarEngraseCilindros", _
        "Error: " & Err.Description
End Sub

Public Sub CargarPurgadores()
    On Error GoTo ErrHandler
    
    Set m_purgadoresIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Purgadores")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_purgadoresCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_purgadoresArray(1 To m_purgadoresCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_purgadoresArray(i).PURG_ID = NzLong(rs("PURG_ID"))
        m_purgadoresArray(i).PURG_MODELO = NzVBA(rs("PURG_MODELO"))
        m_purgadoresArray(i).PURG_DESCRIPCION = NzVBA(rs("PURG_DESCRIPCION"))
        m_purgadoresArray(i).PURG_ADVERTENCIA = NzVBA(rs("PURG_ADVERTENCIA"))
        m_purgadoresArray(i).PURG_PRE_COSTE = NzDouble(rs("PURG_PRE_COSTE"))
        
        ' Agregar al índice
        m_purgadoresIndice.Add m_purgadoresArray(i).PURG_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarPurgadores", _
        "Error: " & Err.Description
End Sub

Public Sub CargarResistenciasCalefaccion()
    On Error GoTo ErrHandler
    
    Set m_resistenciasCalefaccionIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM ResistenciasCalefaccion")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_resistenciasCalefaccionCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_resistenciasCalefaccionArray(1 To m_resistenciasCalefaccionCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_resistenciasCalefaccionArray(i).RESC_ID = NzLong(rs("RESC_ID"))
        m_resistenciasCalefaccionArray(i).RESC_MODELO = NzVBA(rs("RESC_MODELO"))
        m_resistenciasCalefaccionArray(i).RESC_DESCRIPCION = NzVBA(rs("RESC_DESCRIPCION"))
        m_resistenciasCalefaccionArray(i).RESC_ADVERTENCIA = NzVBA(rs("RESC_ADVERTENCIA"))
        m_resistenciasCalefaccionArray(i).RESC_PRE_COSTE = NzDouble(rs("RESC_PRE_COSTE"))
        
        ' Agregar al índice
        m_resistenciasCalefaccionIndice.Add m_resistenciasCalefaccionArray(i).RESC_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarResistenciasCalefaccion", _
        "Error: " & Err.Description
End Sub

Public Sub CargarValvulasReguladoras()
    On Error GoTo ErrHandler
    
    Set m_valvulasReguladorasIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM ValvulasReguladoras")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_valvulasReguladorasCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_valvulasReguladorasArray(1 To m_valvulasReguladorasCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_valvulasReguladorasArray(i).VARG_ID = NzLong(rs("VARG_ID"))
        m_valvulasReguladorasArray(i).VARG_MODELO = NzVBA(rs("VARG_MODELO"))
        m_valvulasReguladorasArray(i).VARG_DESCRIPCION = NzVBA(rs("VARG_DESCRIPCION"))
        m_valvulasReguladorasArray(i).VARG_ADVERTENCIA = NzVBA(rs("VARG_ADVERTENCIA"))
        m_valvulasReguladorasArray(i).VARG_PRE_COSTE = NzDouble(rs("VARG_PRE_COSTE"))
        
        ' Agregar al índice
        m_valvulasReguladorasIndice.Add m_valvulasReguladorasArray(i).VARG_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarValvulasReguladoras", _
        "Error: " & Err.Description
End Sub

Public Sub CargarPaises()
    On Error GoTo ErrHandler
    
    Set m_paisesIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Paises")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_paisesCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_paisesArray(1 To m_paisesCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_paisesArray(i).pais_id = NzLong(rs("pais_id"))
        m_paisesArray(i).pais_nombre = NzVBA(rs("pais_nombre"))
        m_paisesArray(i).pais_iso3166_1n = NzVBA(rs("pais_iso3166_1n"))
        m_paisesArray(i).pais_iso3166_1an2 = NzVBA(rs("pais_iso3166_1an2"))
        m_paisesArray(i).pais_iso3166_1an3 = NzVBA(rs("pais_iso3166_1an3"))
        
        ' Agregar al índice
        m_paisesIndice.Add m_paisesArray(i).pais_id, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarPaises", _
        "Error: " & Err.Description
End Sub

Public Sub CargarComerciales()
    On Error GoTo ErrHandler
    
    Set m_comercialesIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Comerciales")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_comercialesCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_comercialesArray(1 To m_comercialesCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_comercialesArray(i).COME_ID = NzLong(rs("COME_ID"))
        m_comercialesArray(i).COME_NOMBRE = NzVBA(rs("COME_NOMBRE"))
        
        ' Agregar al índice
        m_comercialesIndice.Add m_comercialesArray(i).COME_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarComerciales", _
        "Error: " & Err.Description
End Sub

Public Sub CargarOfertantes()
    On Error GoTo ErrHandler
    
    Set m_ofertantesIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Ofertantes")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_ofertantesCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_ofertantesArray(1 To m_ofertantesCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_ofertantesArray(i).OFTN_ID = NzLong(rs("OFTN_ID"))
        m_ofertantesArray(i).OFTN_NOMBRE = NzVBA(rs("OFTN_NOMBRE"))
        
        ' Agregar al índice
        m_ofertantesIndice.Add m_ofertantesArray(i).OFTN_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarOfertantes", _
        "Error: " & Err.Description
End Sub

Public Sub CargarArrancadoresFuerza()
    On Error GoTo ErrHandler
    
    Set m_arrancadoresFuerzaIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM ArrancadoresFuerza")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_arrancadoresFuerzaCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_arrancadoresFuerzaArray(1 To m_arrancadoresFuerzaCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_arrancadoresFuerzaArray(i).ARR1_ID = NzLong(rs("ARR1_ID"))
        m_arrancadoresFuerzaArray(i).ARR1_MODELO = NzVBA(rs("ARR1_MODELO"))
        m_arrancadoresFuerzaArray(i).ARR1_DESCRIPCION = NzVBA(rs("ARR1_DESCRIPCION"))
        m_arrancadoresFuerzaArray(i).ARR1_ADVERTENCIA = NzVBA(rs("ARR1_ADVERTENCIA"))
        m_arrancadoresFuerzaArray(i).ARR1_PRE_COSTE = NzDouble(rs("ARR1_PRE_COSTE"))
        
        ' Agregar al índice
        m_arrancadoresFuerzaIndice.Add m_arrancadoresFuerzaArray(i).ARR1_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarArrancadoresFuerza", _
        "Error: " & Err.Description
End Sub

Public Sub CargarArrancadoresFuerzaControl()
    On Error GoTo ErrHandler
    
    Set m_arrancadoresFuerzaControlIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM ArrancadoresFuerzaControl")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_arrancadoresFuerzaControlCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_arrancadoresFuerzaControlArray(1 To m_arrancadoresFuerzaControlCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_arrancadoresFuerzaControlArray(i).ARR2_ID = NzLong(rs("ARR2_ID"))
        m_arrancadoresFuerzaControlArray(i).ARR2_MODELO = NzVBA(rs("ARR2_MODELO"))
        m_arrancadoresFuerzaControlArray(i).ARR2_DESCRIPCION = NzVBA(rs("ARR2_DESCRIPCION"))
        m_arrancadoresFuerzaControlArray(i).ARR2_ADVERTENCIA = NzVBA(rs("ARR2_ADVERTENCIA"))
        m_arrancadoresFuerzaControlArray(i).ARR2_PRE_COSTE = NzDouble(rs("ARR2_PRE_COSTE"))
        
        ' Agregar al índice
        m_arrancadoresFuerzaControlIndice.Add m_arrancadoresFuerzaControlArray(i).ARR2_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarArrancadoresFuerzaControl", _
        "Error: " & Err.Description
End Sub

Public Sub CargarArrancadoresControl()
    On Error GoTo ErrHandler
    
    Set m_arrancadoresControlIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM ArrancadoresControl")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_arrancadoresControlCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_arrancadoresControlArray(1 To m_arrancadoresControlCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_arrancadoresControlArray(i).ARR3_ID = NzLong(rs("ARR3_ID"))
        m_arrancadoresControlArray(i).ARR3_MODELO = NzVBA(rs("ARR3_MODELO"))
        m_arrancadoresControlArray(i).ARR3_DESCRIPCION = NzVBA(rs("ARR3_DESCRIPCION"))
        m_arrancadoresControlArray(i).ARR3_ADVERTENCIA = NzVBA(rs("ARR3_ADVERTENCIA"))
        m_arrancadoresControlArray(i).ARR3_PRE_COSTE = NzDouble(rs("ARR3_PRE_COSTE"))
        
        ' Agregar al índice
        m_arrancadoresControlIndice.Add m_arrancadoresControlArray(i).ARR3_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarArrancadoresControl", _
        "Error: " & Err.Description
End Sub

Public Sub CargarArrancadoresFuerzaControlPET()
    On Error GoTo ErrHandler
    
    Set m_arrancadoresFuerzaControlPETIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM ArrancadoresFuerzaControlPET")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_arrancadoresFuerzaControlPETCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_arrancadoresFuerzaControlPETArray(1 To m_arrancadoresFuerzaControlPETCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_arrancadoresFuerzaControlPETArray(i).ARR4_ID = NzLong(rs("ARR4_ID"))
        m_arrancadoresFuerzaControlPETArray(i).ARR4_MODELO = NzVBA(rs("ARR4_MODELO"))
        m_arrancadoresFuerzaControlPETArray(i).ARR4_DESCRIPCION = NzVBA(rs("ARR4_DESCRIPCION"))
        m_arrancadoresFuerzaControlPETArray(i).ARR4_ADVERTENCIA = NzVBA(rs("ARR4_ADVERTENCIA"))
        m_arrancadoresFuerzaControlPETArray(i).ARR4_PRE_COSTE = NzDouble(rs("ARR4_PRE_COSTE"))
        
        ' Agregar al índice
        m_arrancadoresFuerzaControlPETIndice.Add m_arrancadoresFuerzaControlPETArray(i).ARR4_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarArrancadoresFuerzaControlPET", _
        "Error: " & Err.Description
End Sub

Public Sub CargarTransmisoresTemperatura()
    On Error GoTo ErrHandler
    
    Set m_transmisoresTemperaturaIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM TransmisoresTemperatura")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_transmisoresTemperaturaCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_transmisoresTemperaturaArray(1 To m_transmisoresTemperaturaCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_transmisoresTemperaturaArray(i).TRAT_ID = NzLong(rs("TRAT_ID"))
        m_transmisoresTemperaturaArray(i).TRAT_MODELO = NzVBA(rs("TRAT_MODELO"))
        m_transmisoresTemperaturaArray(i).TRAT_DESCRIPCION = NzVBA(rs("TRAT_DESCRIPCION"))
        m_transmisoresTemperaturaArray(i).TRAT_ADVERTENCIA = NzVBA(rs("TRAT_ADVERTENCIA"))
        m_transmisoresTemperaturaArray(i).TRAT_PRE_COSTE = NzDouble(rs("TRAT_PRE_COSTE"))
        
        ' Agregar al índice
        m_transmisoresTemperaturaIndice.Add m_transmisoresTemperaturaArray(i).TRAT_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarTransmisoresTemperatura", _
        "Error: " & Err.Description
End Sub

Public Sub CargarTransmisoresPresion()
    On Error GoTo ErrHandler
    
    Set m_transmisoresPresionIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM TransmisoresPresion")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_transmisoresPresionCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_transmisoresPresionArray(1 To m_transmisoresPresionCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_transmisoresPresionArray(i).TRAP_ID = NzLong(rs("TRAP_ID"))
        m_transmisoresPresionArray(i).TRAP_MODELO = NzVBA(rs("TRAP_MODELO"))
        m_transmisoresPresionArray(i).TRAP_DESCRIPCION = NzVBA(rs("TRAP_DESCRIPCION"))
        m_transmisoresPresionArray(i).TRAP_ADVERTENCIA = NzVBA(rs("TRAP_ADVERTENCIA"))
        m_transmisoresPresionArray(i).TRAP_PRE_COSTE = NzDouble(rs("TRAP_PRE_COSTE"))
        
        ' Agregar al índice
        m_transmisoresPresionIndice.Add m_transmisoresPresionArray(i).TRAP_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarTransmisoresPresion", _
        "Error: " & Err.Description
End Sub

Public Sub CargarTermometros()
    On Error GoTo ErrHandler
    
    Set m_termometrosIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Termometros")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_termometrosCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_termometrosArray(1 To m_termometrosCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_termometrosArray(i).TERM_ID = NzLong(rs("TERM_ID"))
        m_termometrosArray(i).TERM_MODELO = NzVBA(rs("TERM_MODELO"))
        m_termometrosArray(i).TERM_DESCRIPCION = NzVBA(rs("TERM_DESCRIPCION"))
        m_termometrosArray(i).TERM_ADVERTENCIA = NzVBA(rs("TERM_ADVERTENCIA"))
        m_termometrosArray(i).TERM_PRE_COSTE = NzDouble(rs("TERM_PRE_COSTE"))
        
        ' Agregar al índice
        m_termometrosIndice.Add m_termometrosArray(i).TERM_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarTermometros", _
        "Error: " & Err.Description
End Sub

Public Sub CargarElectrovalvulasRegulacion()
    On Error GoTo ErrHandler
    
    Set m_electrovalvulasRegulacionIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM ElectrovalvulasRegulacion")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_electrovalvulasRegulacionCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_electrovalvulasRegulacionArray(1 To m_electrovalvulasRegulacionCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_electrovalvulasRegulacionArray(i).ELER_ID = NzLong(rs("ELER_ID"))
        m_electrovalvulasRegulacionArray(i).ELER_MODELO = NzVBA(rs("ELER_MODELO"))
        m_electrovalvulasRegulacionArray(i).ELER_DESCRIPCION = NzVBA(rs("ELER_DESCRIPCION"))
        m_electrovalvulasRegulacionArray(i).ELER_ADVERTENCIA = NzVBA(rs("ELER_ADVERTENCIA"))
        m_electrovalvulasRegulacionArray(i).ELER_PRE_COSTE = NzDouble(rs("ELER_PRE_COSTE"))
        m_electrovalvulasRegulacionArray(i).Campo1 = NzLong(rs("Campo1"))
        
        ' Agregar al índice
        m_electrovalvulasRegulacionIndice.Add m_electrovalvulasRegulacionArray(i).ELER_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarElectrovalvulasRegulacion", _
        "Error: " & Err.Description
End Sub

Public Sub CargarSensoresCaidaVastago()
    On Error GoTo ErrHandler
    
    Set m_sensoresCaidaVastagoIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM SensoresCaidaVastago")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_sensoresCaidaVastagoCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_sensoresCaidaVastagoArray(1 To m_sensoresCaidaVastagoCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_sensoresCaidaVastagoArray(i).SECV_ID = NzLong(rs("SECV_ID"))
        m_sensoresCaidaVastagoArray(i).SECV_MODELO = NzVBA(rs("SECV_MODELO"))
        m_sensoresCaidaVastagoArray(i).SECV_DESCRIPCION = NzVBA(rs("SECV_DESCRIPCION"))
        m_sensoresCaidaVastagoArray(i).SECV_ADVERTENCIA = NzVBA(rs("SECV_ADVERTENCIA"))
        m_sensoresCaidaVastagoArray(i).SECV_PRE_COSTE = NzDouble(rs("SECV_PRE_COSTE"))
        
        ' Agregar al índice
        m_sensoresCaidaVastagoIndice.Add m_sensoresCaidaVastagoArray(i).SECV_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarSensoresCaidaVastago", _
        "Error: " & Err.Description
End Sub

Public Sub CargarInterruptoresVibracion()
    On Error GoTo ErrHandler
    
    Set m_interruptoresVibracionIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM InterruptoresVibracion")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_interruptoresVibracionCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_interruptoresVibracionArray(1 To m_interruptoresVibracionCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_interruptoresVibracionArray(i).INTV_ID = NzLong(rs("INTV_ID"))
        m_interruptoresVibracionArray(i).INTV_MODELO = NzVBA(rs("INTV_MODELO"))
        m_interruptoresVibracionArray(i).INTV_DESCRIPCION = NzVBA(rs("INTV_DESCRIPCION"))
        m_interruptoresVibracionArray(i).INTV_ADVERTENCIA = NzVBA(rs("INTV_ADVERTENCIA"))
        m_interruptoresVibracionArray(i).INTV_PRE_COSTE = NzDouble(rs("INTV_PRE_COSTE"))
        
        ' Agregar al índice
        m_interruptoresVibracionIndice.Add m_interruptoresVibracionArray(i).INTV_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarInterruptoresVibracion", _
        "Error: " & Err.Description
End Sub

Public Sub CargarInterruptoresNivelAceite()
    On Error GoTo ErrHandler
    
    Set m_interruptoresNivelAceiteIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM InterruptoresNivelAceite")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_interruptoresNivelAceiteCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_interruptoresNivelAceiteArray(1 To m_interruptoresNivelAceiteCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_interruptoresNivelAceiteArray(i).INTA_ID = NzLong(rs("INTA_ID"))
        m_interruptoresNivelAceiteArray(i).INTA_MODELO = NzVBA(rs("INTA_MODELO"))
        m_interruptoresNivelAceiteArray(i).INTA_DESCRIPCION = NzVBA(rs("INTA_DESCRIPCION"))
        m_interruptoresNivelAceiteArray(i).INTA_ADVERTENCIA = NzVBA(rs("INTA_ADVERTENCIA"))
        m_interruptoresNivelAceiteArray(i).INTA_PRE_COSTE = NzDouble(rs("INTA_PRE_COSTE"))
        
        ' Agregar al índice
        m_interruptoresNivelAceiteIndice.Add m_interruptoresNivelAceiteArray(i).INTA_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarInterruptoresNivelAceite", _
        "Error: " & Err.Description
End Sub

Public Sub CargarNivelesCondensados()
    On Error GoTo ErrHandler
    
    Set m_nivelesCondensadosIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM NivelesCondensados")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_nivelesCondensadosCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_nivelesCondensadosArray(1 To m_nivelesCondensadosCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_nivelesCondensadosArray(i).NIVC_ID = NzLong(rs("NIVC_ID"))
        m_nivelesCondensadosArray(i).NIVC_MODELO = NzVBA(rs("NIVC_MODELO"))
        m_nivelesCondensadosArray(i).NIVC_DESCRIPCION = NzVBA(rs("NIVC_DESCRIPCION"))
        m_nivelesCondensadosArray(i).NIVC_ADVERTENCIA = NzVBA(rs("NIVC_ADVERTENCIA"))
        m_nivelesCondensadosArray(i).NIVC_PRE_COSTE = NzDouble(rs("NIVC_PRE_COSTE"))
        
        ' Agregar al índice
        m_nivelesCondensadosIndice.Add m_nivelesCondensadosArray(i).NIVC_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarNivelesCondensados", _
        "Error: " & Err.Description
End Sub

Public Sub CargarValvulasTermostaticas()
    On Error GoTo ErrHandler
    
    Set m_valvulasTermostaticasIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM ValvulasTermostaticas")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_valvulasTermostaticasCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_valvulasTermostaticasArray(1 To m_valvulasTermostaticasCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_valvulasTermostaticasArray(i).VATE_ID = NzLong(rs("VATE_ID"))
        m_valvulasTermostaticasArray(i).VATE_MODELO = NzVBA(rs("VATE_MODELO"))
        m_valvulasTermostaticasArray(i).VATE_DESCRIPCION = NzVBA(rs("VATE_DESCRIPCION"))
        m_valvulasTermostaticasArray(i).VATE_ADVERTENCIA = NzVBA(rs("VATE_ADVERTENCIA"))
        m_valvulasTermostaticasArray(i).VATE_PRE_COSTE = NzDouble(rs("VATE_PRE_COSTE"))
        
        ' Agregar al índice
        m_valvulasTermostaticasIndice.Add m_valvulasTermostaticasArray(i).VATE_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarValvulasTermostaticas", _
        "Error: " & Err.Description
End Sub

Public Sub CargarInstrumentaciones()
    On Error GoTo ErrHandler
    
    Set m_instrumentacionesIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Instrumentaciones")
    
    ' Contar registros
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_instrumentacionesCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    
    ' Redimensionar array
    ReDim m_instrumentacionesArray(1 To m_instrumentacionesCount)
    
    ' Cargar datos
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_instrumentacionesArray(i).INTR_ID = NzLong(rs("INTR_ID"))
        m_instrumentacionesArray(i).INTR_MODELO = NzVBA(rs("INTR_MODELO"))
        m_instrumentacionesArray(i).INTR_DESCRIPCION = NzVBA(rs("INTR_DESCRIPCION"))
        m_instrumentacionesArray(i).INTR_ADVERTENCIA = NzVBA(rs("INTR_ADVERTENCIA"))
        m_instrumentacionesArray(i).TRAT_ID = NzLong(rs("TRAT_ID"))
        m_instrumentacionesArray(i).INTR_TRAT_CANTIDAD = NzLong(rs("INTR_TRAT_CANTIDAD"))
        m_instrumentacionesArray(i).TRAP_ID = NzLong(rs("TRAP_ID"))
        m_instrumentacionesArray(i).INTR_TRAP_CANTIDAD = NzLong(rs("INTR_TRAP_CANTIDAD"))
        m_instrumentacionesArray(i).TERM_ID = NzLong(rs("TERM_ID"))
        m_instrumentacionesArray(i).INTR_TERM_CANTIDAD = NzLong(rs("INTR_TERM_CANTIDAD"))
        m_instrumentacionesArray(i).MANO_ID = NzLong(rs("MANO_ID"))
        m_instrumentacionesArray(i).INTR_MANO_CANTIDAD = NzLong(rs("INTR_MANO_CANTIDAD"))
        m_instrumentacionesArray(i).ELER_ID = NzLong(rs("ELER_ID"))
        m_instrumentacionesArray(i).INTR_ELER_CANTIDAD = NzLong(rs("INTR_ELER_CANTIDAD"))
        m_instrumentacionesArray(i).SECV_ID = NzLong(rs("SECV_ID"))
        m_instrumentacionesArray(i).INTR_SECV_CANTIDAD = NzLong(rs("INTR_SECV_CANTIDAD"))
        m_instrumentacionesArray(i).INTV_ID = NzLong(rs("INTV_ID"))
        m_instrumentacionesArray(i).INTR_INTV_CANTIDAD = NzLong(rs("INTR_INTV_CANTIDAD"))
        m_instrumentacionesArray(i).INTA_ID = NzLong(rs("INTA_ID"))
        m_instrumentacionesArray(i).INTR_INTA_CANTIDAD = NzLong(rs("INTR_INTA_CANTIDAD"))
        m_instrumentacionesArray(i).NIVC_ID = NzLong(rs("NIVC_ID"))
        m_instrumentacionesArray(i).INTR_NIVC_CANTIDAD = NzLong(rs("INTR_NIVC_CANTIDAD"))
        m_instrumentacionesArray(i).VATE_ID = NzLong(rs("VATE_ID"))
        m_instrumentacionesArray(i).INTR_VATE_CANTIDAD = NzLong(rs("INTR_VATE_CANTIDAD"))
        
        ' Agregar al índice
        m_instrumentacionesIndice.Add m_instrumentacionesArray(i).INTR_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarInstrumentaciones", _
        "Error: " & Err.Description
End Sub
Public Sub CargarTiposMaterial()
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM TiposMaterial")
    
    If rs.EOF Then Exit Sub
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    
    m_tiposMaterialCount = rs.RecordCount
    ReDim m_tiposMaterialArray(1 To m_tiposMaterialCount)
    Set m_tiposMaterialIndice = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    i = 1
    
    Do While Not rs.EOF
        m_tiposMaterialArray(i).TIMA_ID = NzLong(rs("TIMA_ID"))
        m_tiposMaterialArray(i).TIMA_MATERIAL = NzVBA(rs("TIMA_MATERIAL"))
        
        m_tiposMaterialIndice(m_tiposMaterialArray(i).TIMA_ID) = i
        i = i + 1
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

Public Sub CargarCalderinesPredeterminados()
    On Error GoTo ErrHandler
    
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM CalderinesPredeterminados")
    
    If rs.EOF Then Exit Sub
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    
    m_calderinesPredeterminadosCount = rs.RecordCount
    ReDim m_calderinesPredeterminadosArray(1 To m_calderinesPredeterminadosCount)
    Set m_calderinesPredeterminadosIndice = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    i = 1
    
    Do While Not rs.EOF
        m_calderinesPredeterminadosArray(i).CALP_ID = NzLong(rs("CALP_ID"))
        m_calderinesPredeterminadosArray(i).CALP_MODELO = NzVBA(rs("CALP_MODELO"))
        m_calderinesPredeterminadosArray(i).CVA1_ID = NzLong(rs("CVA1_ID"))
        m_calderinesPredeterminadosArray(i).CLA1_ID = NzLong(rs("CLA1_ID"))
        m_calderinesPredeterminadosArray(i).CALP_CLA1_CANTIDAD = NzLong(rs("CALP_CLA1_CANTIDAD"))
        m_calderinesPredeterminadosArray(i).CVE1_ID = NzLong(rs("CVE1_ID"))
        m_calderinesPredeterminadosArray(i).CLE1_ID = NzLong(rs("CLE1_ID"))
        m_calderinesPredeterminadosArray(i).CALP_CLE1_CANTIDAD = NzLong(rs("CALP_CLE1_CANTIDAD"))
        m_calderinesPredeterminadosArray(i).CVA2_ID = NzLong(rs("CVA2_ID"))
        m_calderinesPredeterminadosArray(i).CLA2_ID = NzLong(rs("CLA2_ID"))
        m_calderinesPredeterminadosArray(i).CALP_CLA2_CANTIDAD = NzLong(rs("CALP_CLA2_CANTIDAD"))
        m_calderinesPredeterminadosArray(i).CVE2_ID = NzLong(rs("CVE2_ID"))
        m_calderinesPredeterminadosArray(i).CLE2_ID = NzLong(rs("CLE2_ID"))
        m_calderinesPredeterminadosArray(i).CALP_CLE2_CANTIDAD = NzLong(rs("CALP_CLE2_CANTIDAD"))
        m_calderinesPredeterminadosArray(i).CVA3_ID = NzLong(rs("CVA3_ID"))
        m_calderinesPredeterminadosArray(i).CLA3_ID = NzLong(rs("CLA3_ID"))
        m_calderinesPredeterminadosArray(i).CALP_CLA3_CANTIDAD = NzLong(rs("CALP_CLA3_CANTIDAD"))
        m_calderinesPredeterminadosArray(i).CVE3_ID = NzLong(rs("CVE3_ID"))
        m_calderinesPredeterminadosArray(i).CLE3_ID = NzLong(rs("CLE3_ID"))
        m_calderinesPredeterminadosArray(i).CALP_CLE3_CANTIDAD = NzLong(rs("CALP_CLE3_CANTIDAD"))
        m_calderinesPredeterminadosArray(i).CVA4_ID = NzLong(rs("CVA4_ID"))
        m_calderinesPredeterminadosArray(i).CLA4_ID = NzLong(rs("CLA4_ID"))
        m_calderinesPredeterminadosArray(i).CALP_CLA4_CANTIDAD = NzLong(rs("CALP_CLA4_CANTIDAD"))
        m_calderinesPredeterminadosArray(i).CVE4_ID = NzLong(rs("CVE4_ID"))
        m_calderinesPredeterminadosArray(i).CLE4_ID = NzLong(rs("CLE4_ID"))
        m_calderinesPredeterminadosArray(i).CALP_CLE4_CANTIDAD = NzLong(rs("CALP_CLE4_CANTIDAD"))
        m_calderinesPredeterminadosArray(i).CVA5_ID = NzLong(rs("CVA5_ID"))
        m_calderinesPredeterminadosArray(i).CLA5_ID = NzLong(rs("CLA5_ID"))
        m_calderinesPredeterminadosArray(i).CALP_CLA5_CANTIDAD = NzLong(rs("CALP_CLA5_CANTIDAD"))
        m_calderinesPredeterminadosArray(i).CVE5_ID = NzLong(rs("CVE5_ID"))
        m_calderinesPredeterminadosArray(i).CLE5_ID = NzLong(rs("CLE5_ID"))
        m_calderinesPredeterminadosArray(i).CALP_CLE5_CANTIDAD = NzLong(rs("CALP_CLE5_CANTIDAD"))
        m_calderinesPredeterminadosArray(i).CVA6_ID = NzLong(rs("CVA6_ID"))
        m_calderinesPredeterminadosArray(i).CLA6_ID = NzLong(rs("CLA6_ID"))
        m_calderinesPredeterminadosArray(i).CALP_CLA6_CANTIDAD = NzLong(rs("CALP_CLA6_CANTIDAD"))
        m_calderinesPredeterminadosArray(i).CVE6_ID = NzLong(rs("CVE6_ID"))
        m_calderinesPredeterminadosArray(i).CLE6_ID = NzLong(rs("CLE6_ID"))
        m_calderinesPredeterminadosArray(i).CALP_CLE6_CANTIDAD = NzLong(rs("CALP_CLE6_CANTIDAD"))
        
        m_calderinesPredeterminadosIndice(m_calderinesPredeterminadosArray(i).CALP_ID) = i
        i = i + 1
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarCalderinesPredeterminados", "Error: " & Err.Description
End Sub

Public Sub CargarManoObraCabezales()
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM ManoObraCabezales")
    
    If rs.EOF Then Exit Sub
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    
    m_manoObraCabezalesCount = rs.RecordCount
    ReDim m_manoObraCabezalesArray(1 To m_manoObraCabezalesCount)
    Set m_manoObraCabezalesIndice = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    i = 1
    
    Do While Not rs.EOF
        m_manoObraCabezalesArray(i).MAOC_ID = NzLong(rs("MAOC_ID"))
        m_manoObraCabezalesArray(i).MAOC_MODELO = NzVBA(rs("MAOC_MODELO"))
        m_manoObraCabezalesArray(i).MAOC_FASE1_HORAS = NzDouble(rs("MAOC_FASE1_HORAS"))
        m_manoObraCabezalesArray(i).MAOC_FASE1_PRECIO = NzDouble(rs("MAOC_FASE1_PRECIO"))
        m_manoObraCabezalesArray(i).MAOC_FASE2_HORAS = NzDouble(rs("MAOC_FASE2_HORAS"))
        m_manoObraCabezalesArray(i).MAOC_FASE2_PRECIO = NzDouble(rs("MAOC_FASE2_PRECIO"))
        m_manoObraCabezalesArray(i).MAOC_SOLDA_HORAS = NzDouble(rs("MAOC_SOLDA_HORAS"))
        m_manoObraCabezalesArray(i).MAOC_SOLDA_PRECIO = NzDouble(rs("MAOC_SOLDA_PRECIO"))
        m_manoObraCabezalesArray(i).MAOC_PROBA_HORAS = NzDouble(rs("MAOC_PROBA_HORAS"))
        m_manoObraCabezalesArray(i).MAOC_PROBA_PRECIO = NzDouble(rs("MAOC_PROBA_PRECIO"))
        m_manoObraCabezalesArray(i).MAOC_PINTU_HORAS = NzDouble(rs("MAOC_PINTU_HORAS"))
        m_manoObraCabezalesArray(i).MAOC_PINTU_PRECIO = NzDouble(rs("MAOC_PINTU_PRECIO"))
        m_manoObraCabezalesArray(i).MAOC_ELECT_HORAS = NzDouble(rs("MAOC_ELECT_HORAS"))
        m_manoObraCabezalesArray(i).MAOC_ELECT_PRECIO = NzDouble(rs("MAOC_ELECT_PRECIO"))
        
        m_manoObraCabezalesIndice(m_manoObraCabezalesArray(i).MAOC_ID) = i
        i = i + 1
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

Public Sub CargarManoObraIngenierias()
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM ManoObraIngenierias")
    
    If rs.EOF Then Exit Sub
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    
    m_manoObraIngenieriasCount = rs.RecordCount
    ReDim m_manoObraIngenieriasArray(1 To m_manoObraIngenieriasCount)
    Set m_manoObraIngenieriasIndice = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    i = 1
    
    Do While Not rs.EOF
        m_manoObraIngenieriasArray(i).MAOI_ID = NzLong(rs("MAOI_ID"))
        m_manoObraIngenieriasArray(i).MAOI_MODELO = NzVBA(rs("MAOI_MODELO"))
        m_manoObraIngenieriasArray(i).MAOI_INGEN_HORAS = NzDouble(rs("MAOI_INGEN_HORAS"))
        m_manoObraIngenieriasArray(i).MAOI_INGEN_PRECIO = NzDouble(rs("MAOI_INGEN_PRECIO"))
        
        m_manoObraIngenieriasIndice(m_manoObraIngenieriasArray(i).MAOI_ID) = i
        i = i + 1
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

Public Sub CargarAreas()
    On Error GoTo ErrHandler
    
    Set m_AreasIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Areas")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_AreasCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_AreasArray(1 To m_AreasCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_AreasArray(i).AREA_ID = NzLong(rs("AREA_ID"))
        m_AreasArray(i).AREA_DENOMINACION = NzLong(rs("AREA_DENOMINACION"))
        m_AreasArray(i).AREA_GENE_NOMBRE = NzLong(rs("AREA_GENE_NOMBRE"))
        m_AreasArray(i).AREA_NUMERO_CAMPO = NzLong(rs("AREA_NUMERO_CAMPO"))
        
        m_AreasIndice.Add m_AreasArray(i).AREA_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarAreas", "Error: " & Err.Description
End Sub

Public Sub CargarCalderinesCostes()
    On Error GoTo ErrHandler
    
    Set m_CalderinesCostesIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM CalderinesCostes")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_CalderinesCostesCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_CalderinesCostesArray(1 To m_CalderinesCostesCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_CalderinesCostesArray(i).CACO_ID = NzLong(rs("CACO_ID"))
        m_CalderinesCostesArray(i).CALD_ID = NzLong(rs("CALD_ID"))
        m_CalderinesCostesArray(i).VOLU_ID = NzLong(rs("VOLU_ID"))
        m_CalderinesCostesArray(i).CACO_PRE_COSTE = NzDouble(rs("CACO_PRE_COSTE"))
        m_CalderinesCostesArray(i).CACO_DESCRIPCION = NzDouble(rs("CACO_DESCRIPCION"))
        
        m_CalderinesCostesIndice.Add m_CalderinesCostesArray(i).CACO_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarCalderinesCostes", "Error: " & Err.Description
End Sub

Public Sub CargarDepositosEntrada()
    On Error GoTo ErrHandler
    
    Set m_DepositosEntradaIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM DepositosEntrada")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_DepositosEntradaCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_DepositosEntradaArray(1 To m_DepositosEntradaCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_DepositosEntradaArray(i).DEPE_ID = NzLong(rs("DEPE_ID"))
        m_DepositosEntradaArray(i).DEPE_MODELO = NzVBA(rs("DEPE_MODELO"))
        m_DepositosEntradaArray(i).DEPE_DESCRIPCION = NzVBA(rs("DEPE_DESCRIPCION"))
        m_DepositosEntradaArray(i).DEPE_ADVERTENCIA = NzVBA(rs("DEPE_ADVERTENCIA"))
        m_DepositosEntradaArray(i).DEPE_PRE_COSTE = NzDouble(rs("DEPE_PRE_COSTE"))
        
        m_DepositosEntradaIndice.Add m_DepositosEntradaArray(i).DEPE_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarDepositosEntrada", "Error: " & Err.Description
End Sub

Public Sub CargarDepositosFinal()
    On Error GoTo ErrHandler
    
    Set m_DepositosFinalIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM DepositosFinal")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_DepositosFinalCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_DepositosFinalArray(1 To m_DepositosFinalCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_DepositosFinalArray(i).DEPF_ID = NzLong(rs("DEPF_ID"))
        m_DepositosFinalArray(i).DEPF_MODELO = NzVBA(rs("DEPF_MODELO"))
        m_DepositosFinalArray(i).DEPF_DESCRIPCION = NzVBA(rs("DEPF_DESCRIPCION"))
        m_DepositosFinalArray(i).DEPF_ADVERTENCIA = NzVBA(rs("DEPF_ADVERTENCIA"))
        m_DepositosFinalArray(i).DEPF_PRE_COSTE = NzDouble(rs("DEPF_PRE_COSTE"))
        
        m_DepositosFinalIndice.Add m_DepositosFinalArray(i).DEPF_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarDepositosFinal", "Error: " & Err.Description
End Sub

Public Sub CargarMotoresClass()
    On Error GoTo ErrHandler
    
    Set m_MotoresClassIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM MotoresClass")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_MotoresClassCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_MotoresClassArray(1 To m_MotoresClassCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_MotoresClassArray(i).MTCL_ID = NzLong(rs("MTCL_ID"))
        m_MotoresClassArray(i).MTCL_NOMBRE = NzVBA(rs("MTCL_NOMBRE"))
        
        m_MotoresClassIndice.Add m_MotoresClassArray(i).MTCL_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarMotoresClass", "Error: " & Err.Description
End Sub

Public Sub CargarMotoresEspeciales()
    On Error GoTo ErrHandler
    
    Set m_MotoresEspecialesIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM MotoresEspeciales")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_MotoresEspecialesCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_MotoresEspecialesArray(1 To m_MotoresEspecialesCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_MotoresEspecialesArray(i).MTSP_ID = NzLong(rs("MTSP_ID"))
        m_MotoresEspecialesArray(i).MTSP_MODELO = NzVBA(rs("MTSP_MODELO"))
        m_MotoresEspecialesArray(i).MTSP_PRE_COSTE = NzDouble(rs("MTSP_PRE_COSTE"))
        
        m_MotoresEspecialesIndice.Add m_MotoresEspecialesArray(i).MTSP_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarMotoresEspeciales", "Error: " & Err.Description
End Sub

Public Sub CargarMotoresFabrica()
    On Error GoTo ErrHandler
    
    Set m_MotoresFabricaIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM MotoresFabrica")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_MotoresFabricaCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_MotoresFabricaArray(1 To m_MotoresFabricaCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_MotoresFabricaArray(i).MTFB_ID = NzLong(rs("MTFB_ID"))
        m_MotoresFabricaArray(i).MTFB_NOMBRE = NzVBA(rs("MTFB_NOMBRE"))
        
        m_MotoresFabricaIndice.Add m_MotoresFabricaArray(i).MTFB_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarMotoresFabrica", "Error: " & Err.Description
End Sub

Public Sub CargarMotoresNormativas()
    On Error GoTo ErrHandler
    
    Set m_MotoresNormativasIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM MotoresNormativas")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_MotoresNormativasCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_MotoresNormativasArray(1 To m_MotoresNormativasCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_MotoresNormativasArray(i).MTNO_ID = NzLong(rs("MTNO_ID"))
        m_MotoresNormativasArray(i).MTNO_NOMBRE = NzVBA(rs("MTNO_NOMBRE"))
        
        m_MotoresNormativasIndice.Add m_MotoresNormativasArray(i).MTNO_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarMotoresNormativas", "Error: " & Err.Description
End Sub

Public Sub CargarMotoresTension()
    On Error GoTo ErrHandler
    
    Set m_MotoresTensionIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM MotoresTension")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_MotoresTensionCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_MotoresTensionArray(1 To m_MotoresTensionCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_MotoresTensionArray(i).MTTN_ID = NzLong(rs("MTTN_ID"))
        m_MotoresTensionArray(i).MTTN_NOMBRE = NzVBA(rs("MTTN_NOMBRE"))
        
        m_MotoresTensionIndice.Add m_MotoresTensionArray(i).MTTN_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarMotoresTension", "Error: " & Err.Description
End Sub

Public Sub CargarParametros()
    On Error GoTo ErrHandler
    
    Set m_ParametrosIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Parametros")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_ParametrosCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_ParametrosArray(1 To m_ParametrosCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_ParametrosArray(i).PARA_ID = NzLong(rs("PARA_ID"))
        m_ParametrosArray(i).PARA_CALD_EXTRA = NzDouble(rs("PARA_CALD_EXTRA"))
        m_ParametrosArray(i).PARA_CALD_RADIO = NzLong(rs("PARA_CALD_RADIO"))
        m_ParametrosArray(i).PARA_CALD_ASME = NzDouble(rs("PARA_CALD_ASME"))
        m_ParametrosArray(i).PARA_CALD_INOX = NzDouble(rs("PARA_CALD_INOX"))
        m_ParametrosArray(i).PARA_CALD_SELLO = NzLong(rs("PARA_CALD_SELLO"))
        m_ParametrosArray(i).PARA_FEC_ULT_MODIF = NzDate(rs("PARA_FEC_ULT_MODIF"))
        
        m_ParametrosIndice.Add m_ParametrosArray(i).PARA_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarParametros", "Error: " & Err.Description
End Sub

Public Sub CargarRefrigeradoresCostes()
    On Error GoTo ErrHandler
    
    Set m_RefrigeradoresCostesIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM RefrigeradoresCostes")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_RefrigeradoresCostesCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_RefrigeradoresCostesArray(1 To m_RefrigeradoresCostesCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_RefrigeradoresCostesArray(i).RECO_ID = NzLong(rs("RECO_ID"))
        m_RefrigeradoresCostesArray(i).REFR_ID = NzLong(rs("REFR_ID"))
        m_RefrigeradoresCostesArray(i).TIPN_ID = NzLong(rs("TIPN_ID"))
        m_RefrigeradoresCostesArray(i).RECO_INOX_PRE_COSTE = NzDouble(rs("RECO_INOX_PRE_COSTE"))
        m_RefrigeradoresCostesArray(i).RECO_COBRE_PRE_COSTE = NzDouble(rs("RECO_COBRE_PRE_COSTE"))
        m_RefrigeradoresCostesArray(i).RECO_CENT_PRE_COSTE = NzDouble(rs("RECO_CENT_PRE_COSTE"))
        m_RefrigeradoresCostesArray(i).RECO_CSAL_PRE_COSTE = NzDouble(rs("RECO_CSAL_PRE_COSTE"))
        
        m_RefrigeradoresCostesIndice.Add m_RefrigeradoresCostesArray(i).RECO_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarRefrigeradoresCostes", "Error: " & Err.Description
End Sub

Public Sub CargarRefrigeradoresEspeciales()
    On Error GoTo ErrHandler
    
    Set m_RefrigeradoresEspecialesIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM RefrigeradoresEspeciales")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_RefrigeradoresEspecialesCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_RefrigeradoresEspecialesArray(1 To m_RefrigeradoresEspecialesCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_RefrigeradoresEspecialesArray(i).REES_ID = NzLong(rs("REES_ID"))
        m_RefrigeradoresEspecialesArray(i).REES_MODELO = NzVBA(rs("REES_MODELO"))
        m_RefrigeradoresEspecialesArray(i).REES_DESCRIPCION = NzVBA(rs("REES_DESCRIPCION"))
        m_RefrigeradoresEspecialesArray(i).REES_ADVERTENCIA = NzVBA(rs("REES_ADVERTENCIA"))
        m_RefrigeradoresEspecialesArray(i).REES_PRE_COSTE = NzDouble(rs("REES_PRE_COSTE"))
        
        m_RefrigeradoresEspecialesIndice.Add m_RefrigeradoresEspecialesArray(i).REES_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarRefrigeradoresEspeciales", "Error: " & Err.Description
End Sub

Public Sub CargarTiposPresionNominal()
    On Error GoTo ErrHandler
    
    Set m_TiposPresionNominalIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM TiposPresionNominal")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_TiposPresionNominalCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_TiposPresionNominalArray(1 To m_TiposPresionNominalCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_TiposPresionNominalArray(i).TIPN_ID = NzLong(rs("TIPN_ID"))
        m_TiposPresionNominalArray(i).TIPN_PRESION = NzLong(rs("TIPN_PRESION"))
        
        m_TiposPresionNominalIndice.Add m_TiposPresionNominalArray(i).TIPN_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarTiposPresionNominal", "Error: " & Err.Description
End Sub

Public Sub CargarTiposTubos()
    On Error GoTo ErrHandler
    
    Set m_TiposTubosIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM TiposTubos")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_TiposTubosCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_TiposTubosArray(1 To m_TiposTubosCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_TiposTubosArray(i).TITU_ID = NzLong(rs("TITU_ID"))
        m_TiposTubosArray(i).TITU_MATERIAL = NzVBA(rs("TITU_MATERIAL"))
        
        m_TiposTubosIndice.Add m_TiposTubosArray(i).TITU_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarTiposTubos", "Error: " & Err.Description
End Sub

Public Sub CargarTuberiasAgua()
    On Error GoTo ErrHandler
    
    Set m_TuberiasAguaIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM TuberiasAgua")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_TuberiasAguaCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_TuberiasAguaArray(1 To m_TuberiasAguaCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_TuberiasAguaArray(i).TUAG_ID = NzLong(rs("TUAG_ID"))
        m_TuberiasAguaArray(i).TUAG_MODELO = NzVBA(rs("TUAG_MODELO"))
        m_TuberiasAguaArray(i).TUAG_DESCRIPCION = NzVBA(rs("TUAG_DESCRIPCION"))
        m_TuberiasAguaArray(i).TUAG_ADVERTENCIA = NzVBA(rs("TUAG_ADVERTENCIA"))
        m_TuberiasAguaArray(i).TUAG_PRE_COSTE = NzDouble(rs("TUAG_PRE_COSTE"))
        
        m_TuberiasAguaIndice.Add m_TuberiasAguaArray(i).TUAG_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarTuberiasAgua", "Error: " & Err.Description
End Sub

Public Sub CargarTuberiasAire()
    On Error GoTo ErrHandler
    
    Set m_TuberiasAireIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM TuberiasAire")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_TuberiasAireCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_TuberiasAireArray(1 To m_TuberiasAireCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_TuberiasAireArray(i).TUAI_ID = NzLong(rs("TUAI_ID"))
        m_TuberiasAireArray(i).TUAI_MODELO = NzVBA(rs("TUAI_MODELO"))
        
        m_TuberiasAireIndice.Add m_TuberiasAireArray(i).TUAI_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarTuberiasAire", "Error: " & Err.Description
End Sub

Public Sub CargarTuberiasAireDiametros()
    On Error GoTo ErrHandler
    
    Set m_TuberiasAireDiametrosIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM TuberiasAireDiametros")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_TuberiasAireDiametrosCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_TuberiasAireDiametrosArray(1 To m_TuberiasAireDiametrosCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_TuberiasAireDiametrosArray(i).TUDI_ID = NzLong(rs("TUDI_ID"))
        m_TuberiasAireDiametrosArray(i).TUAI_ID = NzLong(rs("TUAI_ID"))
        m_TuberiasAireDiametrosArray(i).TUDI_DIAMETRO = NzDouble(rs("TUDI_DIAMETRO"))
        m_TuberiasAireDiametrosArray(i).TUDI_MODELO = NzVBA(rs("TUDI_MODELO"))
        m_TuberiasAireDiametrosArray(i).TUDI_DESCRIPCION = NzVBA(rs("TUDI_DESCRIPCION"))
        m_TuberiasAireDiametrosArray(i).TUDI_ADVERTENCIA = NzVBA(rs("TUDI_ADVERTENCIA"))
        m_TuberiasAireDiametrosArray(i).TUDI_PRE_COSTE = NzDouble(rs("TUDI_PRE_COSTE"))
        
        m_TuberiasAireDiametrosIndice.Add m_TuberiasAireDiametrosArray(i).TUDI_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarTuberiasAireDiametros", "Error: " & Err.Description
End Sub

Public Sub CargarValvulasSeguridad()
    On Error GoTo ErrHandler
    
    Set m_ValvulasSeguridadIndice = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM ValvulasSeguridad")
    
    If Not rs.EOF Then
        rs.MoveLast
        rs.MoveFirst
    End If
    m_ValvulasSeguridadCount = rs.RecordCount
    If rs.RecordCount = 0 Then Exit Sub
    ReDim m_ValvulasSeguridadArray(1 To m_ValvulasSeguridadCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_ValvulasSeguridadArray(i).VASE_ID = NzLong(rs("VASE_ID"))
        m_ValvulasSeguridadArray(i).VASE_MODELO = NzVBA(rs("VASE_MODELO"))
        m_ValvulasSeguridadArray(i).VASE_DESCRIPCION = NzVBA(rs("VASE_DESCRIPCION"))
        m_ValvulasSeguridadArray(i).VASE_ADVERTENCIA = NzVBA(rs("VASE_ADVERTENCIA"))
        m_ValvulasSeguridadArray(i).VASE_PRE_COSTE = NzDouble(rs("VASE_PRE_COSTE"))
        
        m_ValvulasSeguridadIndice.Add m_ValvulasSeguridadArray(i).VASE_ID, i
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarValvulasSeguridad", "Error: " & Err.Description
End Sub



'=========================================================
' MÉTODOS PÚBLICOS OBTENER - CATÁLOGOS ORIGINALES (15)
'=========================================================

Public Function ObtenerChasis(ByVal CHAS_ID As Long) As tCatalogChasis
    If CHAS_ID <= 0 Then Exit Function
    If Not m_chasisIndice.Exists(CHAS_ID) Then Exit Function
    Dim idx As Long
    idx = m_chasisIndice(CHAS_ID)
    ObtenerChasis = m_chasisArray(idx)
End Function

Public Function ObtenerCorreas(ByVal CORR_ID As Long) As tCatalogCorreas
    If CORR_ID <= 0 Then Exit Function
    If Not m_correasIndice.Exists(CORR_ID) Then Exit Function
    Dim idx As Long
    idx = m_correasIndice(CORR_ID)
    ObtenerCorreas = m_correasArray(idx)
End Function

Public Function ObtenerCabezales(ByVal CABE_ID As Long) As tCatalogCabezales
    If CABE_ID <= 0 Then Exit Function
    If Not m_cabezalesIndice.Exists(CABE_ID) Then Exit Function
    Dim idx As Long
    idx = m_cabezalesIndice(CABE_ID)
    ObtenerCabezales = m_cabezalesArray(idx)
End Function

Public Function ObtenerCilindros(ByVal CILI_ID As Long) As tCatalogCilindros
    If CILI_ID <= 0 Then Exit Function
    If Not m_cilindrosIndice.Exists(CILI_ID) Then Exit Function
    Dim idx As Long
    idx = m_cilindrosIndice(CILI_ID)
    ObtenerCilindros = m_cilindrosArray(idx)
End Function

Public Function ObtenerMotores(ByVal MOTO_ID As Long) As tCatalogMotores
    If MOTO_ID <= 0 Then Exit Function
    If Not m_motoresIndice.Exists(MOTO_ID) Then Exit Function
    Dim idx As Long
    idx = m_motoresIndice(MOTO_ID)
    ObtenerMotores = m_motoresArray(idx)
End Function

Public Function ObtenerGases(ByVal GASE_ID As Long) As tCatalogGases
    If GASE_ID <= 0 Then Exit Function
    If Not m_gasesIndice.Exists(GASE_ID) Then Exit Function
    Dim idx As Long
    idx = m_gasesIndice(GASE_ID)
    ObtenerGases = m_gasesArray(idx)
End Function

Public Function ObtenerRefrigeradores(ByVal REFR_ID As Long) As tCatalogRefrigeradores
    If REFR_ID <= 0 Then Exit Function
    If Not m_refrigeradoresIndice.Exists(REFR_ID) Then Exit Function
    Dim idx As Long
    idx = m_refrigeradoresIndice(REFR_ID)
    ObtenerRefrigeradores = m_refrigeradoresArray(idx)
End Function

Public Function ObtenerCalderines(ByVal CALD_ID As Long) As tCatalogCalderines
    If CALD_ID <= 0 Then Exit Function
    If Not m_calderinesIndice.Exists(CALD_ID) Then Exit Function
    Dim idx As Long
    idx = m_calderinesIndice(CALD_ID)
    ObtenerCalderines = m_calderinesArray(idx)
End Function

Public Function ObtenerVolumenes(ByVal VOLU_ID As Long) As tCatalogVolumenes
    If VOLU_ID <= 0 Then Exit Function
    If Not m_volumenesIndice.Exists(VOLU_ID) Then Exit Function
    Dim idx As Long
    idx = m_volumenesIndice(VOLU_ID)
    ObtenerVolumenes = m_volumenesArray(idx)
End Function

Public Function ObtenerAcoplamientosDirecto(ByVal ACOD_ID As Long) As tCatalogAcoplamientosDirecto
    If ACOD_ID <= 0 Then Exit Function
    If Not m_acoplamientosDirectoIndice.Exists(ACOD_ID) Then Exit Function
    Dim idx As Long
    idx = m_acoplamientosDirectoIndice(ACOD_ID)
    ObtenerAcoplamientosDirecto = m_acoplamientosDirectoArray(idx)
End Function

Public Function ObtenerReductores(ByVal REDU_ID As Long) As tCatalogReductores
    If REDU_ID <= 0 Then Exit Function
    If Not m_reductoresIndice.Exists(REDU_ID) Then Exit Function
    Dim idx As Long
    idx = m_reductoresIndice(REDU_ID)
    ObtenerReductores = m_reductoresArray(idx)
End Function

Public Function ObtenerValvulasRetencion(ByVal VARE_ID As Long) As tCatalogValvulasRetencion
    If VARE_ID <= 0 Then Exit Function
    If Not m_valvulasRetencionIndice.Exists(VARE_ID) Then Exit Function
    Dim idx As Long
    idx = m_valvulasRetencionIndice(VARE_ID)
    ObtenerValvulasRetencion = m_valvulasRetencionArray(idx)
End Function

Public Function ObtenerElectrovalvulas(ByVal ELEC_ID As Long) As tCatalogElectrovalvulas
    If ELEC_ID <= 0 Then Exit Function
    If Not m_electrovalvulasIndice.Exists(ELEC_ID) Then Exit Function
    Dim idx As Long
    idx = m_electrovalvulasIndice(ELEC_ID)
    ObtenerElectrovalvulas = m_electrovalvulasArray(idx)
End Function

Public Function ObtenerFiltros(ByVal FILT_ID As Long) As tCatalogFiltros
    If FILT_ID <= 0 Then Exit Function
    If Not m_filtrosIndice.Exists(FILT_ID) Then Exit Function
    Dim idx As Long
    idx = m_filtrosIndice(FILT_ID)
    ObtenerFiltros = m_filtrosArray(idx)
End Function

Public Function ObtenerManometros(ByVal MANO_ID As Long) As tCatalogManometros
    If MANO_ID <= 0 Then Exit Function
    If Not m_manometrosIndice.Exists(MANO_ID) Then Exit Function
    Dim idx As Long
    idx = m_manometrosIndice(MANO_ID)
    ObtenerManometros = m_manometrosArray(idx)
End Function


'=========================================================
' MÉTODOS PÚBLICOS OBTENER - CATÁLOGOS FASE 1 (32)
'=========================================================

'@Description: Obtiene un registro de Embalajes por ID
'@Param EMBA_ID: Clave primaria
'@Returns: tCatalogEmbalajes
Public Function ObtenerEmbalajes(ByVal EMBA_ID As Long) As tCatalogEmbalajes
    If EMBA_ID <= 0 Then Exit Function
    If Not m_embalajesIndice.Exists(EMBA_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_embalajesIndice(EMBA_ID)
    ObtenerEmbalajes = m_embalajesArray(idx)
End Function

'@Description: Obtiene un registro de Transportes por ID
'@Param TRAN_ID: Clave primaria
'@Returns: tCatalogTransportes
Public Function ObtenerTransportes(ByVal TRAN_ID As Long) As tCatalogTransportes
    If TRAN_ID <= 0 Then Exit Function
    If Not m_transportesIndice.Exists(TRAN_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_transportesIndice(TRAN_ID)
    ObtenerTransportes = m_transportesArray(idx)
End Function

'@Description: Obtiene un registro de PuestasMarcha por ID
'@Param PUMA_ID: Clave primaria
'@Returns: tCatalogPuestasMarcha
Public Function ObtenerPuestasMarcha(ByVal PUMA_ID As Long) As tCatalogPuestasMarcha
    If PUMA_ID <= 0 Then Exit Function
    If Not m_puestasMarchaIndice.Exists(PUMA_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_puestasMarchaIndice(PUMA_ID)
    ObtenerPuestasMarcha = m_puestasMarchaArray(idx)
End Function

'@Description: Obtiene un registro de PruebasAsistencia por ID
'@Param PRAS_ID: Clave primaria
'@Returns: tCatalogPruebasAsistencia
Public Function ObtenerPruebasAsistencia(ByVal PRAS_ID As Long) As tCatalogPruebasAsistencia
    If PRAS_ID <= 0 Then Exit Function
    If Not m_pruebasAsistenciaIndice.Exists(PRAS_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_pruebasAsistenciaIndice(PRAS_ID)
    ObtenerPruebasAsistencia = m_pruebasAsistenciaArray(idx)
End Function

'@Description: Obtiene un registro de Aeros por ID
'@Param AERO_ID: Clave primaria
'@Returns: tCatalogAeros
Public Function ObtenerAeros(ByVal AERO_ID As Long) As tCatalogAeros
    If AERO_ID <= 0 Then Exit Function
    If Not m_aerosIndice.Exists(AERO_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_aerosIndice(AERO_ID)
    ObtenerAeros = m_aerosArray(idx)
End Function

'@Description: Obtiene un registro de CajaLocal por ID
'@Param CALO_ID: Clave primaria
'@Returns: tCatalogCajaLocal
Public Function ObtenerCajaLocal(ByVal CALO_ID As Long) As tCatalogCajaLocal
    If CALO_ID <= 0 Then Exit Function
    If Not m_cajaLocalIndice.Exists(CALO_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_cajaLocalIndice(CALO_ID)
    ObtenerCajaLocal = m_cajaLocalArray(idx)
End Function

'@Description: Obtiene un registro de GruposEngrase por ID
'@Param GREN_ID: Clave primaria
'@Returns: tCatalogGruposEngrase
Public Function ObtenerGruposEngrase(ByVal GREN_ID As Long) As tCatalogGruposEngrase
    If GREN_ID <= 0 Then Exit Function
    If Not m_gruposEngraseIndice.Exists(GREN_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_gruposEngraseIndice(GREN_ID)
    ObtenerGruposEngrase = m_gruposEngraseArray(idx)
End Function

'@Description: Obtiene un registro de LlavesEntrada por ID
'@Param LLEN_ID: Clave primaria
'@Returns: tCatalogLlavesEntrada
Public Function ObtenerLlavesEntrada(ByVal LLEN_ID As Long) As tCatalogLlavesEntrada
    If LLEN_ID <= 0 Then Exit Function
    If Not m_llavesEntradaIndice.Exists(LLEN_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_llavesEntradaIndice(LLEN_ID)
    ObtenerLlavesEntrada = m_llavesEntradaArray(idx)
End Function

'@Description: Obtiene un registro de LlavesSalida por ID
'@Param LLSA_ID: Clave primaria
'@Returns: tCatalogLlavesSalida
Public Function ObtenerLlavesSalida(ByVal LLSA_ID As Long) As tCatalogLlavesSalida
    If LLSA_ID <= 0 Then Exit Function
    If Not m_llavesSalidaIndice.Exists(LLSA_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_llavesSalidaIndice(LLSA_ID)
    ObtenerLlavesSalida = m_llavesSalidaArray(idx)
End Function

'@Description: Obtiene un registro de NormativasCompresor por ID
'@Param NORC_ID: Clave primaria
'@Returns: tCatalogNormativasCompresor
Public Function ObtenerNormativasCompresor(ByVal NORC_ID As Long) As tCatalogNormativasCompresor
    If NORC_ID <= 0 Then Exit Function
    If Not m_normativasCompresorIndice.Exists(NORC_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_normativasCompresorIndice(NORC_ID)
    ObtenerNormativasCompresor = m_normativasCompresorArray(idx)
End Function

'@Description: Obtiene un registro de ValvulasPresion por ID
'@Param VAPR_ID: Clave primaria
'@Returns: tCatalogValvulasPresion
Public Function ObtenerValvulasPresion(ByVal VAPR_ID As Long) As tCatalogValvulasPresion
    If VAPR_ID <= 0 Then Exit Function
    If Not m_valvulasPresionIndice.Exists(VAPR_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_valvulasPresionIndice(VAPR_ID)
    ObtenerValvulasPresion = m_valvulasPresionArray(idx)
End Function

'@Description: Obtiene un registro de EngraseCilindros por ID
'@Param ENCI_ID: Clave primaria
'@Returns: tCatalogEngraseCilindros
Public Function ObtenerEngraseCilindros(ByVal ENCI_ID As Long) As tCatalogEngraseCilindros
    If ENCI_ID <= 0 Then Exit Function
    If Not m_engraseCilindrosIndice.Exists(ENCI_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_engraseCilindrosIndice(ENCI_ID)
    ObtenerEngraseCilindros = m_engraseCilindrosArray(idx)
End Function

'@Description: Obtiene un registro de Purgadores por ID
'@Param PURG_ID: Clave primaria
'@Returns: tCatalogPurgadores
Public Function ObtenerPurgadores(ByVal PURG_ID As Long) As tCatalogPurgadores
    If PURG_ID <= 0 Then Exit Function
    If Not m_purgadoresIndice.Exists(PURG_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_purgadoresIndice(PURG_ID)
    ObtenerPurgadores = m_purgadoresArray(idx)
End Function

'@Description: Obtiene un registro de ResistenciasCalefaccion por ID
'@Param RESC_ID: Clave primaria
'@Returns: tCatalogResistenciasCalefaccion
Public Function ObtenerResistenciasCalefaccion(ByVal RESC_ID As Long) As tCatalogResistenciasCalefaccion
    If RESC_ID <= 0 Then Exit Function
    If Not m_resistenciasCalefaccionIndice.Exists(RESC_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_resistenciasCalefaccionIndice(RESC_ID)
    ObtenerResistenciasCalefaccion = m_resistenciasCalefaccionArray(idx)
End Function

'@Description: Obtiene un registro de ValvulasReguladoras por ID
'@Param VARG_ID: Clave primaria
'@Returns: tCatalogValvulasReguladoras
Public Function ObtenerValvulasReguladoras(ByVal VARG_ID As Long) As tCatalogValvulasReguladoras
    If VARG_ID <= 0 Then Exit Function
    If Not m_valvulasReguladorasIndice.Exists(VARG_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_valvulasReguladorasIndice(VARG_ID)
    ObtenerValvulasReguladoras = m_valvulasReguladorasArray(idx)
End Function

'@Description: Obtiene un registro de Paises por ID
'@Param pais_id: Clave primaria
'@Returns: tCatalogPaises
Public Function ObtenerPaises(ByVal pais_id As Long) As tCatalogPaises
    If pais_id <= 0 Then Exit Function
    If Not m_paisesIndice.Exists(pais_id) Then Exit Function
    
    Dim idx As Long
    idx = m_paisesIndice(pais_id)
    ObtenerPaises = m_paisesArray(idx)
End Function

'@Description: Obtiene un registro de Comerciales por ID
'@Param COME_ID: Clave primaria
'@Returns: tCatalogComerciales
Public Function ObtenerComerciales(ByVal COME_ID As Long) As tCatalogComerciales
    If COME_ID <= 0 Then Exit Function
    If Not m_comercialesIndice.Exists(COME_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_comercialesIndice(COME_ID)
    ObtenerComerciales = m_comercialesArray(idx)
End Function

'@Description: Obtiene un registro de Ofertantes por ID
'@Param OFTN_ID: Clave primaria
'@Returns: tCatalogOfertantes
Public Function ObtenerOfertantes(ByVal OFTN_ID As Long) As tCatalogOfertantes
    If OFTN_ID <= 0 Then Exit Function
    If Not m_ofertantesIndice.Exists(OFTN_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_ofertantesIndice(OFTN_ID)
    ObtenerOfertantes = m_ofertantesArray(idx)
End Function

'@Description: Obtiene un registro de ArrancadoresFuerza por ID
'@Param ARR1_ID: Clave primaria
'@Returns: tCatalogArrancadoresFuerza
Public Function ObtenerArrancadoresFuerza(ByVal ARR1_ID As Long) As tCatalogArrancadoresFuerza
    If ARR1_ID <= 0 Then Exit Function
    If Not m_arrancadoresFuerzaIndice.Exists(ARR1_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_arrancadoresFuerzaIndice(ARR1_ID)
    ObtenerArrancadoresFuerza = m_arrancadoresFuerzaArray(idx)
End Function

'@Description: Obtiene un registro de ArrancadoresFuerzaControl por ID
'@Param ARR2_ID: Clave primaria
'@Returns: tCatalogArrancadoresFuerzaControl
Public Function ObtenerArrancadoresFuerzaControl(ByVal ARR2_ID As Long) As tCatalogArrancadoresFuerzaControl
    If ARR2_ID <= 0 Then Exit Function
    If Not m_arrancadoresFuerzaControlIndice.Exists(ARR2_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_arrancadoresFuerzaControlIndice(ARR2_ID)
    ObtenerArrancadoresFuerzaControl = m_arrancadoresFuerzaControlArray(idx)
End Function

'@Description: Obtiene un registro de ArrancadoresControl por ID
'@Param ARR3_ID: Clave primaria
'@Returns: tCatalogArrancadoresControl
Public Function ObtenerArrancadoresControl(ByVal ARR3_ID As Long) As tCatalogArrancadoresControl
    If ARR3_ID <= 0 Then Exit Function
    If Not m_arrancadoresControlIndice.Exists(ARR3_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_arrancadoresControlIndice(ARR3_ID)
    ObtenerArrancadoresControl = m_arrancadoresControlArray(idx)
End Function

'@Description: Obtiene un registro de ArrancadoresFuerzaControlPET por ID
'@Param ARR4_ID: Clave primaria
'@Returns: tCatalogArrancadoresFuerzaControlPET
Public Function ObtenerArrancadoresFuerzaControlPET(ByVal ARR4_ID As Long) As tCatalogArrancadoresFuerzaControlPET
    If ARR4_ID <= 0 Then Exit Function
    If Not m_arrancadoresFuerzaControlPETIndice.Exists(ARR4_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_arrancadoresFuerzaControlPETIndice(ARR4_ID)
    ObtenerArrancadoresFuerzaControlPET = m_arrancadoresFuerzaControlPETArray(idx)
End Function

'@Description: Obtiene un registro de TransmisoresTemperatura por ID
'@Param TRAT_ID: Clave primaria
'@Returns: tCatalogTransmisoresTemperatura
Public Function ObtenerTransmisoresTemperatura(ByVal TRAT_ID As Long) As tCatalogTransmisoresTemperatura
    If TRAT_ID <= 0 Then Exit Function
    If Not m_transmisoresTemperaturaIndice.Exists(TRAT_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_transmisoresTemperaturaIndice(TRAT_ID)
    ObtenerTransmisoresTemperatura = m_transmisoresTemperaturaArray(idx)
End Function

'@Description: Obtiene un registro de TransmisoresPresion por ID
'@Param TRAP_ID: Clave primaria
'@Returns: tCatalogTransmisoresPresion
Public Function ObtenerTransmisoresPresion(ByVal TRAP_ID As Long) As tCatalogTransmisoresPresion
    If TRAP_ID <= 0 Then Exit Function
    If Not m_transmisoresPresionIndice.Exists(TRAP_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_transmisoresPresionIndice(TRAP_ID)
    ObtenerTransmisoresPresion = m_transmisoresPresionArray(idx)
End Function

'@Description: Obtiene un registro de Termometros por ID
'@Param TERM_ID: Clave primaria
'@Returns: tCatalogTermometros
Public Function ObtenerTermometros(ByVal TERM_ID As Long) As tCatalogTermometros
    If TERM_ID <= 0 Then Exit Function
    If Not m_termometrosIndice.Exists(TERM_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_termometrosIndice(TERM_ID)
    ObtenerTermometros = m_termometrosArray(idx)
End Function

'@Description: Obtiene un registro de ElectrovalvulasRegulacion por ID
'@Param ELER_ID: Clave primaria
'@Returns: tCatalogElectrovalvulasRegulacion
Public Function ObtenerElectrovalvulasRegulacion(ByVal ELER_ID As Long) As tCatalogElectrovalvulasRegulacion
    If ELER_ID <= 0 Then Exit Function
    If Not m_electrovalvulasRegulacionIndice.Exists(ELER_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_electrovalvulasRegulacionIndice(ELER_ID)
    ObtenerElectrovalvulasRegulacion = m_electrovalvulasRegulacionArray(idx)
End Function

'@Description: Obtiene un registro de SensoresCaidaVastago por ID
'@Param SECV_ID: Clave primaria
'@Returns: tCatalogSensoresCaidaVastago
Public Function ObtenerSensoresCaidaVastago(ByVal SECV_ID As Long) As tCatalogSensoresCaidaVastago
    If SECV_ID <= 0 Then Exit Function
    If Not m_sensoresCaidaVastagoIndice.Exists(SECV_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_sensoresCaidaVastagoIndice(SECV_ID)
    ObtenerSensoresCaidaVastago = m_sensoresCaidaVastagoArray(idx)
End Function

'@Description: Obtiene un registro de InterruptoresVibracion por ID
'@Param INTV_ID: Clave primaria
'@Returns: tCatalogInterruptoresVibracion
Public Function ObtenerInterruptoresVibracion(ByVal INTV_ID As Long) As tCatalogInterruptoresVibracion
    If INTV_ID <= 0 Then Exit Function
    If Not m_interruptoresVibracionIndice.Exists(INTV_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_interruptoresVibracionIndice(INTV_ID)
    ObtenerInterruptoresVibracion = m_interruptoresVibracionArray(idx)
End Function

'@Description: Obtiene un registro de InterruptoresNivelAceite por ID
'@Param INTA_ID: Clave primaria
'@Returns: tCatalogInterruptoresNivelAceite
Public Function ObtenerInterruptoresNivelAceite(ByVal INTA_ID As Long) As tCatalogInterruptoresNivelAceite
    If INTA_ID <= 0 Then Exit Function
    If Not m_interruptoresNivelAceiteIndice.Exists(INTA_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_interruptoresNivelAceiteIndice(INTA_ID)
    ObtenerInterruptoresNivelAceite = m_interruptoresNivelAceiteArray(idx)
End Function

'@Description: Obtiene un registro de NivelesCondensados por ID
'@Param NIVC_ID: Clave primaria
'@Returns: tCatalogNivelesCondensados
Public Function ObtenerNivelesCondensados(ByVal NIVC_ID As Long) As tCatalogNivelesCondensados
    If NIVC_ID <= 0 Then Exit Function
    If Not m_nivelesCondensadosIndice.Exists(NIVC_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_nivelesCondensadosIndice(NIVC_ID)
    ObtenerNivelesCondensados = m_nivelesCondensadosArray(idx)
End Function

'@Description: Obtiene un registro de ValvulasTermostaticas por ID
'@Param VATE_ID: Clave primaria
'@Returns: tCatalogValvulasTermostaticas
Public Function ObtenerValvulasTermostaticas(ByVal VATE_ID As Long) As tCatalogValvulasTermostaticas
    If VATE_ID <= 0 Then Exit Function
    If Not m_valvulasTermostaticasIndice.Exists(VATE_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_valvulasTermostaticasIndice(VATE_ID)
    ObtenerValvulasTermostaticas = m_valvulasTermostaticasArray(idx)
End Function

'@Description: Obtiene un registro de Instrumentaciones por ID
'@Param INTR_ID: Clave primaria
'@Returns: tCatalogInstrumentaciones
Public Function ObtenerInstrumentaciones(ByVal INTR_ID As Long) As tCatalogInstrumentaciones
    If INTR_ID <= 0 Then Exit Function
    If Not m_instrumentacionesIndice.Exists(INTR_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_instrumentacionesIndice(INTR_ID)
    ObtenerInstrumentaciones = m_instrumentacionesArray(idx)
End Function

Public Function ObtenerTiposMaterial(ByVal TIMA_ID As Long) As tCatalogTiposMaterial
    If TIMA_ID <= 0 Then Exit Function
    If Not m_tiposMaterialIndice.Exists(TIMA_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_tiposMaterialIndice(TIMA_ID)
    ObtenerTiposMaterial = m_tiposMaterialArray(idx)
End Function

Public Function ObtenerCalderinesPredeterminados(ByVal CALP_ID As Long) As tCatalogCalderinesPredeterminados
    If CALP_ID <= 0 Then Exit Function
    If Not m_calderinesPredeterminadosIndice.Exists(CALP_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_calderinesPredeterminadosIndice(CALP_ID)
    ObtenerCalderinesPredeterminados = m_calderinesPredeterminadosArray(idx)
End Function

Public Function ObtenerManoObraCabezales(ByVal MAOC_ID As Long) As tCatalogManoObraCabezales
    If MAOC_ID <= 0 Then Exit Function
    If Not m_manoObraCabezalesIndice.Exists(MAOC_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_manoObraCabezalesIndice(MAOC_ID)
    ObtenerManoObraCabezales = m_manoObraCabezalesArray(idx)
End Function

Public Function ObtenerManoObraIngenierias(ByVal MAOI_ID As Long) As tCatalogManoObraIngenierias
    If MAOI_ID <= 0 Then Exit Function
    If Not m_manoObraIngenieriasIndice.Exists(MAOI_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_manoObraIngenieriasIndice(MAOI_ID)
    ObtenerManoObraIngenierias = m_manoObraIngenieriasArray(idx)
End Function
'@Description: Obtiene un registro de Areas por ID
'@Param AREA_ID: Clave primaria
'@Returns: tCatalogAreas
Public Function ObtenerAreas(ByVal AREA_ID As Long) As tCatalogAreas
    If AREA_ID <= 0 Then Exit Function
    If Not m_AreasIndice.Exists(AREA_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_AreasIndice(AREA_ID)
    ObtenerAreas = m_AreasArray(idx)
End Function

'@Description: Obtiene un registro de CalderinesCostes por ID
'@Param CACO_ID: Clave primaria
'@Returns: tCatalogCalderinesCostes
Public Function ObtenerCalderinesCostes(ByVal CACO_ID As Long) As tCatalogCalderinesCostes
    If CACO_ID <= 0 Then Exit Function
    If Not m_CalderinesCostesIndice.Exists(CACO_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_CalderinesCostesIndice(CACO_ID)
    ObtenerCalderinesCostes = m_CalderinesCostesArray(idx)
End Function

'@Description: Obtiene un registro de DepositosEntrada por ID
'@Param DEPE_ID: Clave primaria
'@Returns: tCatalogDepositosEntrada
Public Function ObtenerDepositosEntrada(ByVal DEPE_ID As Long) As tCatalogDepositosEntrada
    If DEPE_ID <= 0 Then Exit Function
    If Not m_DepositosEntradaIndice.Exists(DEPE_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_DepositosEntradaIndice(DEPE_ID)
    ObtenerDepositosEntrada = m_DepositosEntradaArray(idx)
End Function

'@Description: Obtiene un registro de DepositosFinal por ID
'@Param DEPF_ID: Clave primaria
'@Returns: tCatalogDepositosFinal
Public Function ObtenerDepositosFinal(ByVal DEPF_ID As Long) As tCatalogDepositosFinal
    If DEPF_ID <= 0 Then Exit Function
    If Not m_DepositosFinalIndice.Exists(DEPF_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_DepositosFinalIndice(DEPF_ID)
    ObtenerDepositosFinal = m_DepositosFinalArray(idx)
End Function

'@Description: Obtiene un registro de MotoresClass por ID
'@Param MTCL_ID: Clave primaria
'@Returns: tCatalogMotoresClass
Public Function ObtenerMotoresClass(ByVal MTCL_ID As Long) As tCatalogMotoresClass
    If MTCL_ID <= 0 Then Exit Function
    If Not m_MotoresClassIndice.Exists(MTCL_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_MotoresClassIndice(MTCL_ID)
    ObtenerMotoresClass = m_MotoresClassArray(idx)
End Function

'@Description: Obtiene un registro de MotoresEspeciales por ID
'@Param MTSP_ID: Clave primaria
'@Returns: tCatalogMotoresEspeciales
Public Function ObtenerMotoresEspeciales(ByVal MTSP_ID As Long) As tCatalogMotoresEspeciales
    If MTSP_ID <= 0 Then Exit Function
    If Not m_MotoresEspecialesIndice.Exists(MTSP_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_MotoresEspecialesIndice(MTSP_ID)
    ObtenerMotoresEspeciales = m_MotoresEspecialesArray(idx)
End Function

'@Description: Obtiene un registro de MotoresFabrica por ID
'@Param MTFB_ID: Clave primaria
'@Returns: tCatalogMotoresFabrica
Public Function ObtenerMotoresFabrica(ByVal MTFB_ID As Long) As tCatalogMotoresFabrica
    If MTFB_ID <= 0 Then Exit Function
    If Not m_MotoresFabricaIndice.Exists(MTFB_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_MotoresFabricaIndice(MTFB_ID)
    ObtenerMotoresFabrica = m_MotoresFabricaArray(idx)
End Function

'@Description: Obtiene un registro de MotoresNormativas por ID
'@Param MTNO_ID: Clave primaria
'@Returns: tCatalogMotoresNormativas
Public Function ObtenerMotoresNormativas(ByVal MTNO_ID As Long) As tCatalogMotoresNormativas
    If MTNO_ID <= 0 Then Exit Function
    If Not m_MotoresNormativasIndice.Exists(MTNO_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_MotoresNormativasIndice(MTNO_ID)
    ObtenerMotoresNormativas = m_MotoresNormativasArray(idx)
End Function

'@Description: Obtiene un registro de MotoresTension por ID
'@Param MTTN_ID: Clave primaria
'@Returns: tCatalogMotoresTension
Public Function ObtenerMotoresTension(ByVal MTTN_ID As Long) As tCatalogMotoresTension
    If MTTN_ID <= 0 Then Exit Function
    If Not m_MotoresTensionIndice.Exists(MTTN_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_MotoresTensionIndice(MTTN_ID)
    ObtenerMotoresTension = m_MotoresTensionArray(idx)
End Function

'@Description: Obtiene un registro de Parametros por ID
'@Param PARA_ID: Clave primaria
'@Returns: tCatalogParametros
Public Function ObtenerParametros(ByVal PARA_ID As Long) As tCatalogParametros
    If PARA_ID <= 0 Then Exit Function
    If Not m_ParametrosIndice.Exists(PARA_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_ParametrosIndice(PARA_ID)
    ObtenerParametros = m_ParametrosArray(idx)
End Function

'@Description: Obtiene un registro de RefrigeradoresCostes por ID
'@Param RECO_ID: Clave primaria
'@Returns: tCatalogRefrigeradoresCostes
Public Function ObtenerRefrigeradoresCostes(ByVal RECO_ID As Long) As tCatalogRefrigeradoresCostes
    If RECO_ID <= 0 Then Exit Function
    If Not m_RefrigeradoresCostesIndice.Exists(RECO_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_RefrigeradoresCostesIndice(RECO_ID)
    ObtenerRefrigeradoresCostes = m_RefrigeradoresCostesArray(idx)
End Function

'@Description: Obtiene un registro de RefrigeradoresEspeciales por ID
'@Param REES_ID: Clave primaria
'@Returns: tCatalogRefrigeradoresEspeciales
Public Function ObtenerRefrigeradoresEspeciales(ByVal REES_ID As Long) As tCatalogRefrigeradoresEspeciales
    If REES_ID <= 0 Then Exit Function
    If Not m_RefrigeradoresEspecialesIndice.Exists(REES_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_RefrigeradoresEspecialesIndice(REES_ID)
    ObtenerRefrigeradoresEspeciales = m_RefrigeradoresEspecialesArray(idx)
End Function

'@Description: Obtiene un registro de TiposPresionNominal por ID
'@Param TIPN_ID: Clave primaria
'@Returns: tCatalogTiposPresionNominal
Public Function ObtenerTiposPresionNominal(ByVal TIPN_ID As Long) As tCatalogTiposPresionNominal
    If TIPN_ID <= 0 Then Exit Function
    If Not m_TiposPresionNominalIndice.Exists(TIPN_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_TiposPresionNominalIndice(TIPN_ID)
    ObtenerTiposPresionNominal = m_TiposPresionNominalArray(idx)
End Function

'@Description: Obtiene un registro de TiposTubos por ID
'@Param TITU_ID: Clave primaria
'@Returns: tCatalogTiposTubos
Public Function ObtenerTiposTubos(ByVal TITU_ID As Long) As tCatalogTiposTubos
    If TITU_ID <= 0 Then Exit Function
    If Not m_TiposTubosIndice.Exists(TITU_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_TiposTubosIndice(TITU_ID)
    ObtenerTiposTubos = m_TiposTubosArray(idx)
End Function

'@Description: Obtiene un registro de TuberiasAgua por ID
'@Param TUAG_ID: Clave primaria
'@Returns: tCatalogTuberiasAgua
Public Function ObtenerTuberiasAgua(ByVal TUAG_ID As Long) As tCatalogTuberiasAgua
    If TUAG_ID <= 0 Then Exit Function
    If Not m_TuberiasAguaIndice.Exists(TUAG_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_TuberiasAguaIndice(TUAG_ID)
    ObtenerTuberiasAgua = m_TuberiasAguaArray(idx)
End Function

'@Description: Obtiene un registro de TuberiasAire por ID
'@Param TUAI_ID: Clave primaria
'@Returns: tCatalogTuberiasAire
Public Function ObtenerTuberiasAire(ByVal TUAI_ID As Long) As tCatalogTuberiasAire
    If TUAI_ID <= 0 Then Exit Function
    If Not m_TuberiasAireIndice.Exists(TUAI_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_TuberiasAireIndice(TUAI_ID)
    ObtenerTuberiasAire = m_TuberiasAireArray(idx)
End Function

'@Description: Obtiene un registro de TuberiasAireDiametros por ID
'@Param TUDI_ID: Clave primaria
'@Returns: tCatalogTuberiasAireDiametros
Public Function ObtenerTuberiasAireDiametros(ByVal TUDI_ID As Long) As tCatalogTuberiasAireDiametros
    If TUDI_ID <= 0 Then Exit Function
    If Not m_TuberiasAireDiametrosIndice.Exists(TUDI_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_TuberiasAireDiametrosIndice(TUDI_ID)
    ObtenerTuberiasAireDiametros = m_TuberiasAireDiametrosArray(idx)
End Function

'@Description: Obtiene un registro de ValvulasSeguridad por ID
'@Param VASE_ID: Clave primaria
'@Returns: tCatalogValvulasSeguridad
Public Function ObtenerValvulasSeguridad(ByVal VASE_ID As Long) As tCatalogValvulasSeguridad
    If VASE_ID <= 0 Then Exit Function
    If Not m_ValvulasSeguridadIndice.Exists(VASE_ID) Then Exit Function
    
    Dim idx As Long
    idx = m_ValvulasSeguridadIndice(VASE_ID)
    ObtenerValvulasSeguridad = m_ValvulasSeguridadArray(idx)
End Function


