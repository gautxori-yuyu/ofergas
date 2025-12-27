Attribute VB_Name = "modMACROLeerOfertas"
'=========================================================
' MÓDULO DE MACROS PARA PRUEBAS Y OPERACIONES CON OFERTAS
'=========================================================
Option Explicit

Const RUTA_BD As String = "C:\Program Files (x86)\Ofertas_Gas\BaseDatos\Ofertas_Gas.mdb"
Sub Macro1()
    Dim shname
    On Error Resume Next
    For Each shname In Array("Oferta", "OfertasOtros", "Ofertas", "Catalogos", "TestCompleto", "TEST_VOLCADO_COMPLETO")
    With Sheets.Add(After:=ActiveSheet)
        .Name = shname
    End With
    Next
    On Error GoTo 0
End Sub

'=========================================================
' FUNCIÓN AUXILIAR - VALIDACIÓN DE GUID
'=========================================================
Function isGUID(ByVal strGUID As Variant) As Boolean
Attribute isGUID.VB_ProcData.VB_Invoke_Func = " \n0"
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

'=========================================================
'@Description: Lee una oferta y vuelca sus datos generales en Excel
'@Scope: Prueba desde Excel
'@Category: Test
'=========================================================
Public Sub Test_LeerOfertaDatosGenerales()
    On Error GoTo ErrHandler
    
    Const OFER_ID As String = "02AB2383-3451-4837-BB80-1562CFD37F27"  ' ? CAMBIAR POR GUID REAL
    
    Dim ctx As clsDBContext
    Dim cat As clsCatalogos
    Dim repo As clsOfertaRepository
    Dim of As clsOferta
    Dim dg As tOfertasDatosGenerales
    Dim ws As Worksheet
    
    ' Preparar hoja
    Set ws = ThisWorkbook.Worksheets("Oferta")
    ws.Cells.Clear
    
    ' Conectar y cargar catálogos
    Set ctx = New clsDBContext
    ctx.Conectar RUTA_BD
    
    Set cat = New clsCatalogos
    Debug.Print "Cargando catálogos..."
    cat.CargarTodos ctx
    Debug.Print "Catálogos cargados."
    
    ' Configurar repositorio
    Set repo = New clsOfertaRepository
    repo.SetDBContext ctx
    repo.SetCatalogos cat
    
    ' Leer oferta
    Debug.Print "Leyendo oferta " & OFER_ID & "..."
    Set of = repo.LeerPorOferID(OFER_ID)
    dg = of.DatosGenerales
    
    ' Cabeceras
    ws.Range("A1:L1").Value = Array( _
        "OFER_ID", "OFER_NUM_OFERTA", "OFER_FECHA", "OFER_CLIENTE", _
        "OFER_USUARIO_FINAL", "GAS", "OFER_OBSERVACIONES", "OFER_ESTADO_OFERTA", _
        "OFER_CALD_EXTRA", "OFER_CALD_ASME", "OFER_CALD_INOX", "COMPLETA")
    
    ' Datos
    ws.Range("A2").Value = dg.OFER_ID
    ws.Range("B2").Value = dg.OFER_NUM_OFERTA
    ws.Range("C2").Value = dg.OFER_FECHA
    ws.Range("D2").Value = dg.OFER_CLIENTE
    ws.Range("E2").Value = dg.OFER_USUARIO_FINAL
    ws.Range("F2").Value = of.ObtenerNombreGas()
    ws.Range("G2").Value = dg.OFER_OBSERVACIONES
    ws.Range("H2").Value = dg.OFER_ESTADO_OFERTA
    ws.Range("I2").Value = dg.OFER_CALD_EXTRA
    ws.Range("J2").Value = dg.OFER_CALD_ASME
    ws.Range("K2").Value = dg.OFER_CALD_INOX
    ws.Range("L2").Value = IIf(of.EstaCompleta(), "SÍ", "NO")
    
    ' Si no está completa, mostrar qué falta
    If Not of.EstaCompleta() Then
        Dim faltantes As Variant
        faltantes = of.TablasFaltantes()
        ws.Range("M2").Value = "Faltan: " & Join(faltantes, ", ")
    End If
    
    ctx.Desconectar
    
    MsgBox "Oferta cargada correctamente", vbInformation
    Exit Sub
    
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    If Not ctx Is Nothing Then ctx.Desconectar
End Sub

'=========================================================
'@Description: Lee una oferta y vuelca la tabla OfertasOtros en Excel
'@Scope: Prueba desde Excel
'@Category: Test
'=========================================================
Public Sub Test_LeerOfertaConOtros()
    On Error GoTo ErrHandler
    
    Const OFER_ID As String = "02AB2383-3451-4837-BB80-1562CFD37F27"  ' ? CAMBIAR POR GUID REAL
    
    Dim ctx As clsDBContext
    Dim cat As clsCatalogos
    Dim repo As clsOfertaRepository
    Dim of As clsOferta
    Dim ws As Worksheet
    Dim i As Long
    Dim it As clsOfertaOtro
    
    ' Preparar hoja
    Set ws = ThisWorkbook.Worksheets("OfertasOtros")
    ws.Cells.Clear
    
    ' Conectar
    Set ctx = New clsDBContext
    ctx.Conectar RUTA_BD
    
    Set cat = New clsCatalogos
    cat.CargarTodos ctx
    
    Set repo = New clsOfertaRepository
    repo.SetDBContext ctx
    repo.SetCatalogos cat
    
    ' Leer oferta
    Set of = repo.LeerPorOferID(OFER_ID)
    
    ' Cabeceras
    ws.Range("A1:D1").Value = Array( _
        "OFOT_LINEA", "OFOT_DESCRIPCION", "OFOT_PRE_COSTE", "Subtotal")
    
    ' Datos
    For i = 1 To of.Otros.Count
        Set it = of.Otros(i)
        ws.Cells(i + 1, 1).Value = it.OFOT_LINEA
        ws.Cells(i + 1, 2).Value = it.OFOT_DESCRIPCION
        ws.Cells(i + 1, 3).Value = it.OFOT_PRE_COSTE
        ws.Cells(i + 1, 4).Value = it.OFOT_PRE_COSTE
    Next i
    
    ctx.Desconectar
    
    MsgBox "OfertasOtros volcadas: " & of.Otros.Count & " registros", vbInformation
    Exit Sub
    
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    If Not ctx Is Nothing Then ctx.Desconectar
End Sub

'=========================================================
'@Description: Lee todas las ofertas y las vuelca en una hoja
'@Scope: Excel VBA ? Base de datos Access (lectura masiva)
'@Category: Exportación
'=========================================================
Public Sub Test_VolcarTodasLasOfertasAExcel()
    On Error GoTo ErrHandler
    
    Dim ctx As clsDBContext
    Dim cat As clsCatalogos
    Dim repo As clsOfertaRepository
    Dim ofertas As Collection
    Dim of As clsOferta
    Dim dg As tOfertasDatosGenerales
    Dim ws As Worksheet
    Dim fila As Long
    
    ' Preparar hoja destino
    Set ws = ThisWorkbook.Worksheets("Ofertas")
    ws.Cells.Clear
    
    ws.Range("A1:H1").Value = Array( _
        "OFER_ID", "OFER_NUM_OFERTA", "OFER_FECHA", "OFER_CLIENTE", _
        "GAS", "COMPLETA", "TOTAL", "NUM_OTROS")
    
    fila = 2
    
    ' Conectar a base de datos
    Set ctx = New clsDBContext
    ctx.Conectar RUTA_BD
    
    Set cat = New clsCatalogos
    Debug.Print "Cargando catálogos..."
    cat.CargarTodos ctx
    Debug.Print "Catálogos cargados."
    
    Set repo = New clsOfertaRepository
    repo.SetDBContext ctx
    repo.SetCatalogos cat
    
    ' Leer repositorio completo
    Debug.Print "Leyendo todas las ofertas..."
    Set ofertas = repo.LeerTodas()
    Debug.Print "Ofertas leídas: " & ofertas.Count
    
    ' Volcado masivo
    For Each of In ofertas
        dg = of.DatosGenerales
        
        ws.Cells(fila, 1).Value = dg.OFER_ID
        ws.Cells(fila, 2).Value = dg.OFER_NUM_OFERTA
        ws.Cells(fila, 3).Value = dg.OFER_FECHA
        ws.Cells(fila, 4).Value = dg.OFER_CLIENTE
        ws.Cells(fila, 5).Value = of.ObtenerNombreGas()
        ws.Cells(fila, 6).Value = IIf(of.EstaCompleta(), "SÍ", "NO")
        ws.Cells(fila, 7).Value = of.CalcularTotal()
        ws.Cells(fila, 8).Value = of.Otros.Count
        
        fila = fila + 1
    Next of
    
    ' Limpieza
    ctx.Desconectar
    
    MsgBox ofertas.Count & " ofertas volcadas correctamente.", vbInformation
    Exit Sub
    
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    If Not ctx Is Nothing Then ctx.Desconectar
End Sub

'=========================================================
'@Description: Prueba de acceso a catálogos
'@Scope: Test de catálogos
'@Category: Test
'=========================================================
Public Sub Test_Catalogos()
    On Error GoTo ErrHandler
    
    Dim ctx As clsDBContext
    Dim cat As clsCatalogos
    Dim ws As Worksheet
    Dim catChasis As tCatalogChasis
    Dim catGas As tCatalogGases
    
    ' Preparar hoja
    Set ws = ThisWorkbook.Worksheets("Catalogos")
    ws.Cells.Clear
    
    ' Conectar
    Set ctx = New clsDBContext
    ctx.Conectar RUTA_BD
    
    ' Cargar catálogos
    Set cat = New clsCatalogos
    Debug.Print "Cargando catálogos..."
    cat.CargarTodos ctx
    Debug.Print "Catálogos cargados."
    
    ' Probar acceso (cambiar IDs por valores reales de tu BD)
    ws.Range("A1").Value = "Prueba de catálogos:"
    
    ' Chasis ID 1 (ejemplo)
    On Error Resume Next
    catChasis = cat.ObtenerChasis(1)
    If Err.Number = 0 Then
        ws.Range("A3").Value = "Chasis ID 1:"
        ws.Range("B3").Value = catChasis.CHAS_MODELO
        ws.Range("C3").Value = catChasis.CHAS_DESCRIPCION
        ws.Range("D3").Value = catChasis.CHAS_PRE_COSTE
    Else
        ws.Range("A3").Value = "Chasis ID 1 no encontrado"
    End If
    On Error GoTo ErrHandler
    
    ' Gas ID 1 (ejemplo)
    On Error Resume Next
    catGas = cat.ObtenerGases(1)
    If Err.Number = 0 Then
        ws.Range("A4").Value = "Gas ID 1:"
        ws.Range("B4").Value = catGas.GASE_DENOMINACION
    Else
        ws.Range("A4").Value = "Gas ID 1 no encontrado"
    End If
    On Error GoTo ErrHandler
    
    ctx.Desconectar
    
    MsgBox "Prueba de catálogos completada. Revisa la hoja 'Catalogos'", vbInformation
    Exit Sub
    
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    If Not ctx Is Nothing Then ctx.Desconectar
End Sub

'=========================================================
'@Description: Prueba exhaustiva de lectura de oferta completa
'@Scope: Test completo
'@Category: Test
'=========================================================
Public Sub Test_OfertaCompleta()
    On Error GoTo ErrHandler
    
    Const OFER_ID As String = "02AB2383-3451-4837-BB80-1562CFD37F27"  ' ? CAMBIAR POR GUID REAL
    
    Dim ctx As clsDBContext
    Dim cat As clsCatalogos
    Dim repo As clsOfertaRepository
    Dim of As clsOferta
    Dim ws As Worksheet
    Dim fila As Long
    
    ' Preparar hoja
    Set ws = ThisWorkbook.Worksheets("TestCompleto")
    ws.Cells.Clear
    
    ' Conectar
    Set ctx = New clsDBContext
    ctx.Conectar RUTA_BD
    
    Set cat = New clsCatalogos
    cat.CargarTodos ctx
    
    Set repo = New clsOfertaRepository
    repo.SetDBContext ctx
    repo.SetCatalogos cat
    
    ' Leer oferta
    Set of = repo.LeerPorOferID(OFER_ID)
    
    fila = 1
    
    ' Información general
    ws.Cells(fila, 1).Value = "OFERTA COMPLETA - DIAGNÓSTICO"
    ws.Cells(fila, 1).Font.Bold = True
    fila = fila + 2
    
    ws.Cells(fila, 1).Value = "Número de oferta:"
    ws.Cells(fila, 2).Value = of.DatosGenerales.OFER_NUM_OFERTA
    fila = fila + 1
    
    ws.Cells(fila, 1).Value = "Cliente:"
    ws.Cells(fila, 2).Value = of.DatosGenerales.OFER_CLIENTE
    fila = fila + 1
    
    ws.Cells(fila, 1).Value = "Gas:"
    ws.Cells(fila, 2).Value = of.ObtenerNombreGas()
    fila = fila + 1
    
    ws.Cells(fila, 1).Value = "Fecha:"
    ws.Cells(fila, 2).Value = of.DatosGenerales.OFER_FECHA
    fila = fila + 2
    
    ws.Cells(fila, 1).Value = "Oferta completa:"
    ws.Cells(fila, 2).Value = IIf(of.EstaCompleta(), "SÍ", "NO")
    ws.Cells(fila, 2).Font.Bold = True
    If Not of.EstaCompleta() Then
        ws.Cells(fila, 2).Font.Color = RGB(255, 0, 0)
    Else
        ws.Cells(fila, 2).Font.Color = RGB(0, 128, 0)
    End If
    fila = fila + 1
    
    If Not of.EstaCompleta() Then
        Dim faltantes As Variant
        faltantes = of.TablasFaltantes()
        ws.Cells(fila, 1).Value = "Tablas faltantes:"
        ws.Cells(fila, 2).Value = Join(faltantes, ", ")
        ws.Cells(fila, 2).Font.Color = RGB(255, 0, 0)
        fila = fila + 1
    End If
    
    fila = fila + 1
    ws.Cells(fila, 1).Value = "Total calculado:"
    ws.Cells(fila, 2).Value = of.CalcularTotal()
    ws.Cells(fila, 2).NumberFormat = "#,##0.00 €"
    fila = fila + 1
    
    ws.Cells(fila, 1).Value = "Número de 'Otros':"
    ws.Cells(fila, 2).Value = of.Otros.Count
    fila = fila + 2
    
    ' Descripción del chasis (prueba de acceso a catálogo)
    ws.Cells(fila, 1).Value = "Descripción Chasis:"
    ws.Cells(fila, 2).Value = of.ObtenerDescripcionChasis()
    
    ctx.Desconectar
    
    MsgBox "Prueba completa finalizada. Revisa la hoja 'TestCompleto'", vbInformation
    Exit Sub
    
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    If Not ctx Is Nothing Then ctx.Desconectar
End Sub
'=========================================================
' MACRO DE TEST COMPLETO - VOLCADO TOTAL DE OFERTAS
' Crea hoja "TEST_VOLCADO_COMPLETO" con TODOS los campos
' Colecciones en celdas multi-línea con ajuste automático
'=========================================================

Public Sub TEST_VolcadoCompletoOfertas()
    On Error GoTo ErrHandler
    
    Dim ctx As clsDBContext
    Dim repo As clsOfertaRepository
    Dim catalogos As clsCatalogos
    Dim ofertas As Collection
    Dim oferta As clsOferta
    Dim ws As Worksheet
    Dim fila As Long
    Dim col As Long
    
    ' Inicializar
    Set ctx = New clsDBContext
    ctx.Conectar RUTA_BD
    
    Set catalogos = New clsCatalogos
    catalogos.CargarTodos ctx
    
    Set repo = New clsOfertaRepository
    repo.SetDBContext ctx
    repo.SetCatalogos catalogos
    
    ' Cargar todas las ofertas
    Set ofertas = repo.LeerTodas()
    
    ' Crear o limpiar hoja
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("TEST_VOLCADO_COMPLETO")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "TEST_VOLCADO_COMPLETO"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo ErrHandler
    
    ' Configurar hoja
    ws.Cells.Font.Name = "Consolas"
    ws.Cells.Font.Size = 9
    
    ' ENCABEZADOS
    fila = 1
    col = 1
    
    ' Datos Generales (20 columnas)
    ws.Cells(fila, col).Value = "OFER_ID": col = col + 1
    ws.Cells(fila, col).Value = "NUM_OFERTA": col = col + 1
    ws.Cells(fila, col).Value = "FECHA": col = col + 1
    ws.Cells(fila, col).Value = "CLIENTE": col = col + 1
    ws.Cells(fila, col).Value = "USUARIO_FINAL": col = col + 1
    ws.Cells(fila, col).Value = "PAIS": col = col + 1
    ws.Cells(fila, col).Value = "COMERCIAL": col = col + 1
    ws.Cells(fila, col).Value = "OFERTANTE": col = col + 1
    ws.Cells(fila, col).Value = "GAS": col = col + 1
    ws.Cells(fila, col).Value = "NUM_CALCULO": col = col + 1
    ws.Cells(fila, col).Value = "OBSERVACIONES": col = col + 1
    ws.Cells(fila, col).Value = "ESTADO": col = col + 1
    ws.Cells(fila, col).Value = "PRESION_TRABAJO": col = col + 1
    ws.Cells(fila, col).Value = "FEC_ULT_MODIF": col = col + 1
    ws.Cells(fila, col).Value = "CAUDAL": col = col + 1
    ws.Cells(fila, col).Value = "POTENCIA": col = col + 1
    ws.Cells(fila, col).Value = "TENSION": col = col + 1
    ws.Cells(fila, col).Value = "FRECUENCIA": col = col + 1
    ws.Cells(fila, col).Value = "VELOCIDAD_GIRO": col = col + 1
    ws.Cells(fila, col).Value = "TOTAL_OFERTA": col = col + 1
    
    ' Extras (4 columnas)
    ws.Cells(fila, col).Value = "EMBALAJE": col = col + 1
    ws.Cells(fila, col).Value = "TRANSPORTE": col = col + 1
    ws.Cells(fila, col).Value = "PUESTA_MARCHA": col = col + 1
    ws.Cells(fila, col).Value = "PRUEBAS_ASISTENCIA": col = col + 1
    
    ' Motor (10 columnas)
    ws.Cells(fila, col).Value = "MOTOR": col = col + 1
    ws.Cells(fila, col).Value = "CHASIS": col = col + 1
    ws.Cells(fila, col).Value = "CORREA": col = col + 1
    ws.Cells(fila, col).Value = "ACOPLAMIENTO": col = col + 1
    ws.Cells(fila, col).Value = "REDUCTOR": col = col + 1
    ws.Cells(fila, col).Value = "VALVULA_RETENCION": col = col + 1
    ws.Cells(fila, col).Value = "ELECTROVALVULA": col = col + 1
    ws.Cells(fila, col).Value = "FILTRO": col = col + 1
    ws.Cells(fila, col).Value = "MANOMETRO": col = col + 1
    ws.Cells(fila, col).Value = "REFRIGERADOR": col = col + 1
    
    ' Accesorios (6 columnas) - MULTI-LÍNEA
    ws.Cells(fila, col).Value = "ARRANCADORES" & vbLf & "(Multi-línea)": col = col + 1
    ws.Cells(fila, col).Value = "AEROS": col = col + 1
    ws.Cells(fila, col).Value = "CAJA_LOCAL": col = col + 1
    ws.Cells(fila, col).Value = "GRUPOS_ENGRASE": col = col + 1
    ws.Cells(fila, col).Value = "LLAVES_ENTRADA": col = col + 1
    ws.Cells(fila, col).Value = "LLAVES_SALIDA": col = col + 1
    
    ' Cabezal (2 columnas) - MULTI-LÍNEA
    ws.Cells(fila, col).Value = "CABEZAL": col = col + 1
    ws.Cells(fila, col).Value = "NORMATIVA": col = col + 1
    
    ' Opciones (5 columnas)
    ws.Cells(fila, col).Value = "VALVULA_PRESION": col = col + 1
    ws.Cells(fila, col).Value = "ENGRASE_CILINDROS": col = col + 1
    ws.Cells(fila, col).Value = "PURGADOR": col = col + 1
    ws.Cells(fila, col).Value = "RESISTENCIA": col = col + 1
    ws.Cells(fila, col).Value = "VALVULA_REGULADORA": col = col + 1
    
    ' Calderines (1 columna) - MULTI-LÍNEA
    ws.Cells(fila, col).Value = "CALDERINES" & vbLf & "(Multi-línea)": col = col + 1
    
    ' Instrumentación (10 columnas)
    ws.Cells(fila, col).Value = "INSTRUMENTACION": col = col + 1
    ws.Cells(fila, col).Value = "TRANSMISOR_TEMP": col = col + 1
    ws.Cells(fila, col).Value = "TRANSMISOR_PRES": col = col + 1
    ws.Cells(fila, col).Value = "TERMOMETRO": col = col + 1
    ws.Cells(fila, col).Value = "ELECTROVAL_REG": col = col + 1
    ws.Cells(fila, col).Value = "SENSOR_VASTAGO": col = col + 1
    ws.Cells(fila, col).Value = "INTERR_VIBRACION": col = col + 1
    ws.Cells(fila, col).Value = "INTERR_NIVEL_ACEITE": col = col + 1
    ws.Cells(fila, col).Value = "NIVEL_CONDENSADO": col = col + 1
    ws.Cells(fila, col).Value = "VALVULA_TERMOSTATICA": col = col + 1
    
    ' Otros (1 columna) - MULTI-LÍNEA
    ws.Cells(fila, col).Value = "OTROS CONCEPTOS" & vbLf & "(Multi-línea)": col = col + 1
    
    ' Formatear encabezados
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(200, 200, 200)
    ws.Rows(1).WrapText = True
    ws.Rows(1).RowHeight = 30
    
    ' DATOS - Recorrer ofertas
    fila = 2
    Dim i As Long
    For i = 1 To ofertas.Count
        Set oferta = ofertas(i)
        col = 1
        
        ' === DATOS GENERALES ===
        ws.Cells(fila, col).Value = oferta.DatosGenerales.OFER_ID: col = col + 1
        ws.Cells(fila, col).Value = "'" & oferta.DatosGenerales.OFER_NUM_OFERTA: col = col + 1  ' Prefijo ' para texto
        ws.Cells(fila, col).Value = oferta.DatosGenerales.OFER_FECHA: col = col + 1
        ws.Cells(fila, col).Value = oferta.DatosGenerales.OFER_CLIENTE: col = col + 1
        ws.Cells(fila, col).Value = oferta.DatosGenerales.OFER_USUARIO_FINAL: col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombrePais(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreComercial(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreOfertante(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreGas(): col = col + 1
        ws.Cells(fila, col).Value = oferta.DatosGenerales.OFER_NUM_CALCULO: col = col + 1
        ws.Cells(fila, col).Value = oferta.DatosGenerales.OFER_OBSERVACIONES: col = col + 1
        ws.Cells(fila, col).Value = oferta.DatosGenerales.OFER_ESTADO_OFERTA: col = col + 1
        ws.Cells(fila, col).Value = oferta.DatosGenerales.OFER_FEC_ULT_MODIF: col = col + 1
        
'FIXME: El siguiente debería ser un campo calculado
'        ws.Cells(fila, col).Value = oferta.DatosGenerales.OFER_TOTAL_OFERTA: col = col + 1
'FIXME: NO EXISTEN
'        ws.Cells(fila, col).Value = oferta.DatosGenerales.OFER_PRESION_TRABAJO: col = col + 1
'        ws.Cells(fila, col).Value = oferta.DatosGenerales.OFER_CAUDAL: col = col + 1
'        ws.Cells(fila, col).Value = oferta.DatosGenerales.OFER_POTENCIA: col = col + 1
'        ws.Cells(fila, col).Value = oferta.DatosGenerales.OFER_TENSION: col = col + 1
'        ws.Cells(fila, col).Value = oferta.DatosGenerales.OFER_FRECUENCIA: col = col + 1
'        ws.Cells(fila, col).Value = oferta.DatosGenerales.OFER_VELOCIDAD_GIRO: col = col + 1
        
        ' === EXTRAS ===
        ws.Cells(fila, col).Value = oferta.ObtenerNombreEmbalaje(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreTransporte(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombrePuestaMarcha(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombrePruebasAsistencia(): col = col + 1
        
        ' === MOTOR ===
        ws.Cells(fila, col).Value = oferta.ObtenerNombreMotor(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreChasis(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreCorrea(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreAcoplamientoDirecto(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreReductor(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreValvulaRetencion(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreElectrovalvula(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreFiltro(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreManometro(): col = col + 1
'FIXME: Hay que actualizar esta llamada para Obtener la información de todos los refrigeradores
'        ws.Cells(fila, col).Value = oferta.ObtenerNombreRefrigerador(): col = col + 1
        
        ' === ACCESORIOS (MULTI-LÍNEA) ===
        ws.Cells(fila, col).Value = FormatearArrancadores(oferta): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreAero(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreCajaLocal(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreGrupoEngrase(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreLlaveEntrada(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreLlaveSalida(): col = col + 1
        
        ' === CABEZAL (MULTI-LÍNEA) ===
        ws.Cells(fila, col).Value = FormatearCabezal(oferta): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreNormativaCompresor(): col = col + 1
        
        ' === OPCIONES ===
        ws.Cells(fila, col).Value = oferta.ObtenerNombreValvulaPresion(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreEngraseCilindros(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombrePurgador(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreResistenciaCalefaccion(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreValvulaReguladora(): col = col + 1
        
        ' === CALDERINES (MULTI-LÍNEA) ===
        ws.Cells(fila, col).Value = FormatearCalderines(oferta): col = col + 1
        
        ' === INSTRUMENTACIÓN ===
        ws.Cells(fila, col).Value = oferta.ObtenerNombreInstrumentacion(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreTransmisorTemperatura(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreTransmisorPresion(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreTermometro(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreElectrovalvulaRegulacion(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreSensorCaidaVastago(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreInterruptorVibracion(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreInterruptorNivelAceite(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreNivelCondensados(): col = col + 1
        ws.Cells(fila, col).Value = oferta.ObtenerNombreValvulaTermostatica(): col = col + 1
        
        ' === OTROS (MULTI-LÍNEA) ===
        ws.Cells(fila, col).Value = FormatearOtros(oferta): col = col + 1
        
        ' AJUSTAR ALTURA DE FILA para multi-línea
        ws.Rows(fila).WrapText = True
        ws.Rows(fila).AutoFit
        
        fila = fila + 1
    Next i
    
    ' Ajustar anchos de columna
    ws.Columns("A:BZ").AutoFit
    
    ' Inmovilizar paneles (primera fila y primera columna)
    ws.Range("B2").Select
    ActiveWindow.FreezePanes = True
    
    ' Mensaje final
    MsgBox "? Test completado!" & vbCrLf & vbCrLf & _
           "Ofertas procesadas: " & ofertas.Count & vbCrLf & _
           "Hoja: TEST_VOLCADO_COMPLETO", vbInformation
    
    ' Limpiar
    ctx.Desconectar
    Set ctx = Nothing
    Set repo = Nothing
    Set ofertas = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "? Error en TEST_VolcadoCompletoOfertas:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical
    If Not ctx Is Nothing Then ctx.Desconectar
End Sub

'=========================================================
' FUNCIONES AUXILIARES PARA FORMATEAR COLECCIONES
'=========================================================
'FIXME: Esta función está mal implementada
Private Function FormatearArrancadores(ByVal oferta As clsOferta) As String
    Dim resultado As String
    resultado = ""
    
    ' Arrancador 1
    If oferta.Accesorios.ARR1_ID > 0 Then
        resultado = resultado & "ARR1: " & oferta.ObtenerNombreArrancador1() & _
                    " (Cant: " & oferta.Accesorios.OFAC_ARRA1_CANTIDAD & ")"
    End If
    
    ' Arrancador 2
    If oferta.Accesorios.ARR2_ID > 0 Then
        If Len(resultado) > 0 Then resultado = resultado & vbLf
        resultado = resultado & "ARR2: " & oferta.ObtenerNombreArrancador2() & _
                    " (Cant: " & oferta.Accesorios.OFAC_ARRA2_CANTIDAD & ")"
    End If
    
    ' Arrancador 3
    If oferta.Accesorios.ARR3_ID > 0 Then
        If Len(resultado) > 0 Then resultado = resultado & vbLf
        resultado = resultado & "ARR3: " & oferta.ObtenerNombreArrancador3() & _
                    " (Cant: " & oferta.Accesorios.OFAC_ARRA3_CANTIDAD & ")"
    End If
    
    ' Arrancador 4
    If oferta.Accesorios.ARR4_ID > 0 Then
        If Len(resultado) > 0 Then resultado = resultado & vbLf
        resultado = resultado & "ARR4: " & oferta.ObtenerNombreArrancador4() & _
                    " (Cant: " & oferta.Accesorios.OFAC_ARRA4_CANTIDAD & ")"
    End If
    
    FormatearArrancadores = resultado
End Function

Private Function FormatearCabezal(ByVal oferta As clsOferta) As String
    Dim resultado As String
    resultado = ""
    
    ' Cabezal principal
    resultado = oferta.ObtenerNombreCabezal() & vbLf
    resultado = resultado & "Etapas: " & oferta.Cabezal.OFCA_NUM_ETAPAS & vbLf
    
    ' Cilindros por etapa
    If oferta.Cabezal.CIL1_ID > 0 Then
        resultado = resultado & "CIL1: " & oferta.ObtenerNombreCilindro1() & _
                    " (Cant: " & oferta.Cabezal.OFCA_CIL1_CANTIDAD & ")" & vbLf
    End If
    
    If oferta.Cabezal.CIL2_ID > 0 Then
        resultado = resultado & "CIL2: " & oferta.ObtenerNombreCilindro2() & _
                    " (Cant: " & oferta.Cabezal.OFCA_CIL2_CANTIDAD & ")" & vbLf
    End If
    
    If oferta.Cabezal.CIL3_ID > 0 Then
        resultado = resultado & "CIL3: " & oferta.ObtenerNombreCilindro3() & _
                    " (Cant: " & oferta.Cabezal.OFCA_CIL3_CANTIDAD & ")" & vbLf
    End If
    
    If oferta.Cabezal.CIL4_ID > 0 Then
        resultado = resultado & "CIL4: " & oferta.ObtenerNombreCilindro4() & _
                    " (Cant: " & oferta.Cabezal.OFCA_CIL4_CANTIDAD & ")" & vbLf
    End If
    
    If oferta.Cabezal.CIL5_ID > 0 Then
        resultado = resultado & "CIL5: " & oferta.ObtenerNombreCilindro5() & _
                    " (Cant: " & oferta.Cabezal.OFCA_CIL5_CANTIDAD & ")" & vbLf
    End If
    
    If oferta.Cabezal.CIL6_ID > 0 Then
        resultado = resultado & "CIL6: " & oferta.ObtenerNombreCilindro6() & _
                    " (Cant: " & oferta.Cabezal.OFCA_CIL6_CANTIDAD & ")"
    End If
    
    FormatearCabezal = resultado
End Function

Private Function FormatearCalderines(ByVal oferta As clsOferta) As String
    Dim resultado As String
    Dim i As Long
    resultado = ""
    
    ' Recorrer colección de calderines
    For i = 1 To oferta.Calderines.Count
        Dim calderin As tOfertasCalderines
        calderin = oferta.Calderines(i)
        
        If i > 1 Then resultado = resultado & vbLf
        
        resultado = resultado & "Calderín " & i & ": " & _
                    oferta.ObtenerNombreCalderin(i) & _
                    " (Vol: " & calderin.CALD_ID & ")"
    Next i
    
    If Len(resultado) = 0 Then resultado = "(Sin calderines)"
    
    FormatearCalderines = resultado
End Function

Private Function FormatearOtros(ByVal oferta As clsOferta) As String
    Dim resultado As String
    Dim i As Long
    resultado = ""
    
    ' Recorrer colección de OfertasOtros
    For i = 1 To oferta.Otros.Count
        Dim otro As clsOfertaOtro
        Set otro = oferta.Otros(i)
        
        If i > 1 Then resultado = resultado & vbLf
        
        resultado = resultado & otro.OFOT_LINEA & ". " & _
                    otro.OFOT_DESCRIPCION & _
                    " ($" & Format(otro.OFOT_PRE_COSTE, "#,##0.00") & ")"
    Next i
    
    If Len(resultado) = 0 Then resultado = "(Sin otros conceptos)"
    
    FormatearOtros = resultado
End Function

