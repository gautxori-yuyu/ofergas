Attribute VB_Name = "modMACROLeerOfertas"
'=========================================================
' MÓDULO DE MACROS PARA PRUEBAS Y OPERACIONES CON OFERTAS
'=========================================================
Option Explicit

Const RUTA_BD As String = "C:\Program Files (x86)\Ofertas_Gas\BaseDatos\Ofertas_Gas.mdb"

'=========================================================
' FUNCIÓN AUXILIAR - VALIDACIÓN DE GUID
'=========================================================
Function isGUID(ByVal strGUID As Variant) As Boolean
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
    
    Const OFER_ID As String = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"  ' ← CAMBIAR POR GUID REAL
    
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
    
    Const OFER_ID As String = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"  ' ← CAMBIAR POR GUID REAL
    
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
'@Scope: Excel VBA → Base de datos Access (lectura masiva)
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
    catGas = cat.ObtenerGas(1)
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
    
    Const OFER_ID As String = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"  ' ← CAMBIAR POR GUID REAL
    
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
