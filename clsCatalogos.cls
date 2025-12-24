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
'@Purpose: Arrays de UDTs + Dictionary para lookup rápido
'=========================================================
Option Explicit

'=========================================================
' ARRAYS DE UDTs + DICTIONARIES PARA ÍNDICES
'=========================================================
Private m_Chasis() As tCatalogChasis
Private m_ChasisIdx As Object  ' Dictionary: ID -> índice en array
Private m_ChasisCount As Long

Private m_Correas() As tCatalogCorreas
Private m_CorreasIdx As Object
Private m_CorreasCount As Long

Private m_Cabezales() As tCatalogCabezales
Private m_CabezalesIdx As Object
Private m_CabezalesCount As Long

Private m_Cilindros() As tCatalogCilindros
Private m_CilindrosIdx As Object
Private m_CilindrosCount As Long

Private m_Motores() As tCatalogMotores
Private m_MotoresIdx As Object
Private m_MotoresCount As Long

Private m_Gases() As tCatalogGases
Private m_GasesIdx As Object
Private m_GasesCount As Long

Private m_Refrigeradores() As tCatalogRefrigeradores
Private m_RefrigeradoresIdx As Object
Private m_RefrigeradoresCount As Long

Private m_Calderines() As tCatalogCalderines
Private m_CalderinesIdx As Object
Private m_CalderinesCount As Long

Private m_Volumenes() As tCatalogVolumenes
Private m_VolumenesIdx As Object
Private m_VolumenesCount As Long

Private m_AcoplamientosDirecto() As tCatalogAcoplamientosDirecto
Private m_AcoplamientosDirectoIdx As Object
Private m_AcoplamientosDirectoCount As Long

Private m_Reductores() As tCatalogReductores
Private m_ReductoresIdx As Object
Private m_ReductoresCount As Long

Private m_ValvulasRetencion() As tCatalogValvulasRetencion
Private m_ValvulasRetencionIdx As Object
Private m_ValvulasRetencionCount As Long

Private m_Electrovalvulas() As tCatalogElectrovalvulas
Private m_ElectrovalvulasIdx As Object
Private m_ElectrovalvulasCount As Long

Private m_Filtros() As tCatalogFiltros
Private m_FiltrosIdx As Object
Private m_FiltrosCount As Long

Private m_Manometros() As tCatalogManometros
Private m_ManometrosIdx As Object
Private m_ManometrosCount As Long

Private m_Ctx As clsDBContext
Private m_Cargado As Boolean

'=========================================================
' INICIALIZACIÓN Y CARGA
'=========================================================

Private Sub Class_Initialize()
    m_Cargado = False
End Sub

Public Sub CargarTodos(ByVal ctx As clsDBContext)
    On Error GoTo ErrHandler
    
    If m_Cargado Then Exit Sub
    Set m_Ctx = ctx
    
    CargarChasis
    CargarCorreas
    CargarCabezales
    CargarCilindros
    CargarMotores
    CargarGases
    CargarRefrigeradores
    CargarCalderines
    CargarVolumenes
    CargarAcoplamientosDirecto
    CargarReductores
    CargarValvulasRetencion
    CargarElectrovalvulas
    CargarFiltros
    CargarManometros
    
    m_Cargado = True
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "clsCatalogos.CargarTodos", _
        "Error cargando catálogos: " & Err.Description
End Sub

'=========================================================
' MÉTODOS PRIVADOS DE CARGA
'=========================================================

Private Sub CargarChasis()
    On Error GoTo ErrHandler
    
    Set m_ChasisIdx = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Chasis")
    
    ' Contar registros
    rs.MoveLast
    m_ChasisCount = rs.RecordCount
    rs.MoveFirst
    
    ' Redimensionar array
    ReDim m_Chasis(1 To m_ChasisCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_Chasis(i).CHAS_ID = NzLong(rs("CHAS_ID"))
        m_Chasis(i).CHAS_MODELO = NzVBA(rs("CHAS_MODELO"))
        m_Chasis(i).CHAS_DESCRIPCION = NzVBA(rs("CHAS_DESCRIPCION"))
        m_Chasis(i).CHAS_ADVERTENCIA = NzVBA(rs("CHAS_ADVERTENCIA"))
        m_Chasis(i).CHAS_PRE_COSTE = NzDouble(rs("CHAS_PRE_COSTE"))
        
        ' Guardar índice en Dictionary
        m_ChasisIdx.Add m_Chasis(i).CHAS_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarChasis", "Error: " & Err.Description
End Sub

Private Sub CargarCorreas()
    On Error GoTo ErrHandler
    
    Set m_CorreasIdx = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Correas")
    
    rs.MoveLast
    m_CorreasCount = rs.RecordCount
    rs.MoveFirst
    
    ReDim m_Correas(1 To m_CorreasCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_Correas(i).CORR_ID = NzLong(rs("CORR_ID"))
        m_Correas(i).CORR_MODELO = NzVBA(rs("CORR_MODELO"))
        m_Correas(i).CORR_DESCRIPCION = NzVBA(rs("CORR_DESCRIPCION"))
        m_Correas(i).CORR_ADVERTENCIA = NzVBA(rs("CORR_ADVERTENCIA"))
        m_Correas(i).CORR_PRE_COSTE = NzDouble(rs("CORR_PRE_COSTE"))
        m_Correas(i).CORR_ATEX_COSTE = NzDouble(rs("CORR_ATEX_COSTE"))
        
        m_CorreasIdx.Add m_Correas(i).CORR_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarCorreas", "Error: " & Err.Description
End Sub

Private Sub CargarCabezales()
    On Error GoTo ErrHandler
    
    Set m_CabezalesIdx = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Cabezales")
    
    rs.MoveLast
    m_CabezalesCount = rs.RecordCount
    rs.MoveFirst
    
    ReDim m_Cabezales(1 To m_CabezalesCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_Cabezales(i).CABE_ID = NzLong(rs("CABE_ID"))
        m_Cabezales(i).CABE_MODELO = NzVBA(rs("CABE_MODELO"))
        m_Cabezales(i).CABE_DESCRIPCION = NzVBA(rs("CABE_DESCRIPCION"))
        m_Cabezales(i).CABE_ADVERTENCIA = NzVBA(rs("CABE_ADVERTENCIA"))
        m_Cabezales(i).CABE_PRE_COSTE = NzDouble(rs("CABE_PRE_COSTE"))
        m_Cabezales(i).CABE_BLOQUE_SAS = NzVBA(rs("CABE_BLOQUE_SAS"))
        
        m_CabezalesIdx.Add m_Cabezales(i).CABE_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarCabezales", "Error: " & Err.Description
End Sub

Private Sub CargarCilindros()
    On Error GoTo ErrHandler
    
    Set m_CilindrosIdx = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Cilindros")
    
    rs.MoveLast
    m_CilindrosCount = rs.RecordCount
    rs.MoveFirst
    
    ReDim m_Cilindros(1 To m_CilindrosCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_Cilindros(i).CILI_ID = NzLong(rs("CILI_ID"))
        m_Cilindros(i).CILI_MODELO = NzVBA(rs("CILI_MODELO"))
        m_Cilindros(i).CILI_DESCRIPCION = NzVBA(rs("CILI_DESCRIPCION"))
        m_Cilindros(i).CILI_ADVERTENCIA = NzVBA(rs("CILI_ADVERTENCIA"))
        m_Cilindros(i).CILI_PRE_COSTE = NzDouble(rs("CILI_PRE_COSTE"))
        m_Cilindros(i).CILI_T2_COSTE = NzDouble(rs("CILI_T2_COSTE"))
        m_Cilindros(i).CILI_CO2_COSTE = NzDouble(rs("CILI_CO2_COSTE"))
        m_Cilindros(i).CILI_CPI_COSTE = NzDouble(rs("CILI_CPI_COSTE"))
        m_Cilindros(i).CILI_SAS_COSTE = NzDouble(rs("CILI_SAS_COSTE"))
        m_Cilindros(i).CILI_EMP_COSTE = NzDouble(rs("CILI_EMP_COSTE"))
        
        m_CilindrosIdx.Add m_Cilindros(i).CILI_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarCilindros", "Error: " & Err.Description
End Sub

Private Sub CargarMotores()
    On Error GoTo ErrHandler
    
    Set m_MotoresIdx = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Motores")
    
    rs.MoveLast
    m_MotoresCount = rs.RecordCount
    rs.MoveFirst
    
    ReDim m_Motores(1 To m_MotoresCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_Motores(i).MOTO_ID = NzLong(rs("MOTO_ID"))
        m_Motores(i).MOTO_MODELO = NzVBA(rs("MOTO_MODELO"))
        m_Motores(i).MOTO_DESCRIPCION = NzVBA(rs("MOTO_DESCRIPCION"))
        m_Motores(i).MOTO_ADVERTENCIA = NzVBA(rs("MOTO_ADVERTENCIA"))
        m_Motores(i).MOTO_PRE_COSTE = NzDouble(rs("MOTO_PRE_COSTE"))
        m_Motores(i).MTNO_ID = NzLong(rs("MTNO_ID"))
        m_Motores(i).MTCL_ID = NzLong(rs("MTCL_ID"))
        m_Motores(i).MTFB_ID = NzLong(rs("MTFB_ID"))
        m_Motores(i).MTTN_ID = NzLong(rs("MTTN_ID"))
        
        m_MotoresIdx.Add m_Motores(i).MOTO_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarMotores", "Error: " & Err.Description
End Sub

Private Sub CargarGases()
    On Error GoTo ErrHandler
    
    Set m_GasesIdx = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Gases")
    
    rs.MoveLast
    m_GasesCount = rs.RecordCount
    rs.MoveFirst
    
    ReDim m_Gases(1 To m_GasesCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_Gases(i).GASE_ID = NzLong(rs("GASE_ID"))
        m_Gases(i).GASE_DENOMINACION = NzVBA(rs("GASE_DENOMINACION"))
        
        m_GasesIdx.Add m_Gases(i).GASE_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarGases", "Error: " & Err.Description
End Sub

Private Sub CargarRefrigeradores()
    On Error GoTo ErrHandler
    
    Set m_RefrigeradoresIdx = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Refrigeradores")
    
    rs.MoveLast
    m_RefrigeradoresCount = rs.RecordCount
    rs.MoveFirst
    
    ReDim m_Refrigeradores(1 To m_RefrigeradoresCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_Refrigeradores(i).REFR_ID = NzLong(rs("REFR_ID"))
        m_Refrigeradores(i).REFR_MODELO = NzVBA(rs("REFR_MODELO"))
        m_Refrigeradores(i).REFR_DESCRIPCION = NzVBA(rs("REFR_DESCRIPCION"))
        m_Refrigeradores(i).REFR_ADVERTENCIA = NzVBA(rs("REFR_ADVERTENCIA"))
        
        m_RefrigeradoresIdx.Add m_Refrigeradores(i).REFR_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarRefrigeradores", "Error: " & Err.Description
End Sub

Private Sub CargarCalderines()
    On Error GoTo ErrHandler
    
    Set m_CalderinesIdx = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Calderines")
    
    rs.MoveLast
    m_CalderinesCount = rs.RecordCount
    rs.MoveFirst
    
    ReDim m_Calderines(1 To m_CalderinesCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_Calderines(i).CALD_ID = NzLong(rs("CALD_ID"))
        m_Calderines(i).CALD_MODELO = NzVBA(rs("CALD_MODELO"))
        m_Calderines(i).CALD_DESCRIPCION = NzVBA(rs("CALD_DESCRIPCION"))
        m_Calderines(i).CALD_ADVERTENCIA = NzVBA(rs("CALD_ADVERTENCIA"))
        
        m_CalderinesIdx.Add m_Calderines(i).CALD_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarCalderines", "Error: " & Err.Description
End Sub

Private Sub CargarVolumenes()
    On Error GoTo ErrHandler
    
    Set m_VolumenesIdx = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Volumenes")
    
    rs.MoveLast
    m_VolumenesCount = rs.RecordCount
    rs.MoveFirst
    
    ReDim m_Volumenes(1 To m_VolumenesCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_Volumenes(i).VOLU_ID = NzLong(rs("VOLU_ID"))
        m_Volumenes(i).VOLU_CANTIDAD = NzVBA(rs("VOLU_CANTIDAD"))
        
        m_VolumenesIdx.Add m_Volumenes(i).VOLU_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarVolumenes", "Error: " & Err.Description
End Sub

Private Sub CargarAcoplamientosDirecto()
    On Error GoTo ErrHandler
    
    Set m_AcoplamientosDirectoIdx = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM AcoplamientosDirecto")
    
    rs.MoveLast
    m_AcoplamientosDirectoCount = rs.RecordCount
    rs.MoveFirst
    
    ReDim m_AcoplamientosDirecto(1 To m_AcoplamientosDirectoCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_AcoplamientosDirecto(i).ACOD_ID = NzLong(rs("ACOD_ID"))
        m_AcoplamientosDirecto(i).ACOD_MODELO = NzVBA(rs("ACOD_MODELO"))
        m_AcoplamientosDirecto(i).ACOD_DESCRIPCION = NzVBA(rs("ACOD_DESCRIPCION"))
        m_AcoplamientosDirecto(i).ACOD_ADVERTENCIA = NzVBA(rs("ACOD_ADVERTENCIA"))
        m_AcoplamientosDirecto(i).ACOD_PRE_COSTE = NzDouble(rs("ACOD_PRE_COSTE"))
        
        m_AcoplamientosDirectoIdx.Add m_AcoplamientosDirecto(i).ACOD_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarAcoplamientosDirecto", "Error: " & Err.Description
End Sub

Private Sub CargarReductores()
    On Error GoTo ErrHandler
    
    Set m_ReductoresIdx = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Reductores")
    
    rs.MoveLast
    m_ReductoresCount = rs.RecordCount
    rs.MoveFirst
    
    ReDim m_Reductores(1 To m_ReductoresCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_Reductores(i).REDU_ID = NzLong(rs("REDU_ID"))
        m_Reductores(i).REDU_MODELO = NzVBA(rs("REDU_MODELO"))
        m_Reductores(i).REDU_DESCRIPCION = NzVBA(rs("REDU_DESCRIPCION"))
        m_Reductores(i).REDU_ADVERTENCIA = NzVBA(rs("REDU_ADVERTENCIA"))
        m_Reductores(i).REDU_PRE_COSTE = NzDouble(rs("REDU_PRE_COSTE"))
        m_Reductores(i).REDU_ATEX_COSTE = NzDouble(rs("REDU_ATEX_COSTE"))
        
        m_ReductoresIdx.Add m_Reductores(i).REDU_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarReductores", "Error: " & Err.Description
End Sub

Private Sub CargarValvulasRetencion()
    On Error GoTo ErrHandler
    
    Set m_ValvulasRetencionIdx = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM ValvulasRetencion")
    
    rs.MoveLast
    m_ValvulasRetencionCount = rs.RecordCount
    rs.MoveFirst
    
    ReDim m_ValvulasRetencion(1 To m_ValvulasRetencionCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_ValvulasRetencion(i).VARE_ID = NzLong(rs("VARE_ID"))
        m_ValvulasRetencion(i).VARE_MODELO = NzVBA(rs("VARE_MODELO"))
        m_ValvulasRetencion(i).VARE_DESCRIPCION = NzVBA(rs("VARE_DESCRIPCION"))
        m_ValvulasRetencion(i).VARE_ADVERTENCIA = NzVBA(rs("VARE_ADVERTENCIA"))
        m_ValvulasRetencion(i).VARE_PRE_COSTE = NzDouble(rs("VARE_PRE_COSTE"))
        
        m_ValvulasRetencionIdx.Add m_ValvulasRetencion(i).VARE_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarValvulasRetencion", "Error: " & Err.Description
End Sub

Private Sub CargarElectrovalvulas()
    On Error GoTo ErrHandler
    
    Set m_ElectrovalvulasIdx = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Electrovalvulas")
    
    rs.MoveLast
    m_ElectrovalvulasCount = rs.RecordCount
    rs.MoveFirst
    
    ReDim m_Electrovalvulas(1 To m_ElectrovalvulasCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_Electrovalvulas(i).ELEC_ID = NzLong(rs("ELEC_ID"))
        m_Electrovalvulas(i).ELEC_MODELO = NzVBA(rs("ELEC_MODELO"))
        m_Electrovalvulas(i).ELEC_DESCRIPCION = NzVBA(rs("ELEC_DESCRIPCION"))
        m_Electrovalvulas(i).ELEC_ADVERTENCIA = NzVBA(rs("ELEC_ADVERTENCIA"))
        m_Electrovalvulas(i).ELEC_PRE_COSTE = NzDouble(rs("ELEC_PRE_COSTE"))
        
        m_ElectrovalvulasIdx.Add m_Electrovalvulas(i).ELEC_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarElectrovalvulas", "Error: " & Err.Description
End Sub

Private Sub CargarFiltros()
    On Error GoTo ErrHandler
    
    Set m_FiltrosIdx = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Filtros")
    
    rs.MoveLast
    m_FiltrosCount = rs.RecordCount
    rs.MoveFirst
    
    ReDim m_Filtros(1 To m_FiltrosCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_Filtros(i).FILT_ID = NzLong(rs("FILT_ID"))
        m_Filtros(i).FILT_MODELO = NzVBA(rs("FILT_MODELO"))
        m_Filtros(i).FILT_DESCRIPCION = NzVBA(rs("FILT_DESCRIPCION"))
        m_Filtros(i).FILT_ADVERTENCIA = NzVBA(rs("FILT_ADVERTENCIA"))
        m_Filtros(i).FILT_PRE_COSTE = NzDouble(rs("FILT_PRE_COSTE"))
        
        m_FiltrosIdx.Add m_Filtros(i).FILT_ID, i
        
        i = i + 1
        rs.MoveNext
    Loop
    rs.Close
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "clsCatalogos.CargarFiltros", "Error: " & Err.Description
End Sub

Private Sub CargarManometros()
    On Error GoTo ErrHandler
    
    Set m_ManometrosIdx = CreateObject("Scripting.Dictionary")
    Dim rs As Object
    Set rs = m_Ctx.AbrirRecordset("SELECT * FROM Manometros")
    
    rs.MoveLast
    m_ManometrosCount = rs.RecordCount
    rs.MoveFirst
    
    ReDim m_Manometros(1 To m_ManometrosCount)
    
    Dim i As Long
    i = 1
    Do While Not rs.EOF
        m_Manometros(i).MANO_ID = NzLong(rs("MANO_ID"))
        m_Manometros(i).MANO_MODELO = NzVBA(rs("MANO_MODELO"))
        m_Manometros(i).MANO_DESCRIPCION = NzVBA(rs("MANO_DESCRIPCION"))
        m_Manometros(i).MANO_ADVERTENCIA = NzVBA(rs("MANO_ADVERTENCIA"))
        m_Manometros(i).MANO_PRE_COSTE = NzDouble(rs("MANO_PRE_COSTE"))
        
        m_ManometrosIdx.Add m_Manometros(i).MANO_ID, i
        
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
' MÉTODOS PÚBLICOS DE ACCESO
'=========================================================

Public Function ObtenerChasis(ByVal ID As Long) As tCatalogChasis
    On Error GoTo ErrHandler
    If m_ChasisIdx.Exists(ID) Then
        ObtenerChasis = m_Chasis(m_ChasisIdx(ID))
    Else
        Err.Raise 5, , "Chasis ID " & ID & " no encontrado"
    End If
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "clsCatalogos.ObtenerChasis", Err.Description
End Function

Public Function ObtenerCorreas(ByVal ID As Long) As tCatalogCorreas
    On Error GoTo ErrHandler
    If m_CorreasIdx.Exists(ID) Then
        ObtenerCorreas = m_Correas(m_CorreasIdx(ID))
    Else
        Err.Raise 5, , "Correa ID " & ID & " no encontrada"
    End If
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "clsCatalogos.ObtenerCorreas", Err.Description
End Function

Public Function ObtenerCabezal(ByVal ID As Long) As tCatalogCabezales
    On Error GoTo ErrHandler
    If m_CabezalesIdx.Exists(ID) Then
        ObtenerCabezal = m_Cabezales(m_CabezalesIdx(ID))
    Else
        Err.Raise 5, , "Cabezal ID " & ID & " no encontrado"
    End If
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "clsCatalogos.ObtenerCabezal", Err.Description
End Function

Public Function ObtenerCilindro(ByVal ID As Long) As tCatalogCilindros
    On Error GoTo ErrHandler
    If m_CilindrosIdx.Exists(ID) Then
        ObtenerCilindro = m_Cilindros(m_CilindrosIdx(ID))
    Else
        Err.Raise 5, , "Cilindro ID " & ID & " no encontrado"
    End If
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "clsCatalogos.ObtenerCilindro", Err.Description
End Function

Public Function ObtenerMotor(ByVal ID As Long) As tCatalogMotores
    On Error GoTo ErrHandler
    If m_MotoresIdx.Exists(ID) Then
        ObtenerMotor = m_Motores(m_MotoresIdx(ID))
    Else
        Err.Raise 5, , "Motor ID " & ID & " no encontrado"
    End If
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "clsCatalogos.ObtenerMotor", Err.Description
End Function

Public Function ObtenerGas(ByVal ID As Long) As tCatalogGases
    On Error GoTo ErrHandler
    If m_GasesIdx.Exists(ID) Then
        ObtenerGas = m_Gases(m_GasesIdx(ID))
    Else
        Err.Raise 5, , "Gas ID " & ID & " no encontrado"
    End If
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "clsCatalogos.ObtenerGas", Err.Description
End Function

Public Function ObtenerRefrigerador(ByVal ID As Long) As tCatalogRefrigeradores
    On Error GoTo ErrHandler
    If m_RefrigeradoresIdx.Exists(ID) Then
        ObtenerRefrigerador = m_Refrigeradores(m_RefrigeradoresIdx(ID))
    Else
        Err.Raise 5, , "Refrigerador ID " & ID & " no encontrado"
    End If
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "clsCatalogos.ObtenerRefrigerador", Err.Description
End Function

Public Function ObtenerCalderin(ByVal ID As Long) As tCatalogCalderines
    On Error GoTo ErrHandler
    If m_CalderinesIdx.Exists(ID) Then
        ObtenerCalderin = m_Calderines(m_CalderinesIdx(ID))
    Else
        Err.Raise 5, , "Calderín ID " & ID & " no encontrado"
    End If
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "clsCatalogos.ObtenerCalderin", Err.Description
End Function

Public Function ObtenerVolumen(ByVal ID As Long) As tCatalogVolumenes
    On Error GoTo ErrHandler
    If m_VolumenesIdx.Exists(ID) Then
        ObtenerVolumen = m_Volumenes(m_VolumenesIdx(ID))
    Else
        Err.Raise 5, , "Volumen ID " & ID & " no encontrado"
    End If
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "clsCatalogos.ObtenerVolumen", Err.Description
End Function

Public Function ObtenerAcoplamientoDirecto(ByVal ID As Long) As tCatalogAcoplamientosDirecto
    On Error GoTo ErrHandler
    If m_AcoplamientosDirectoIdx.Exists(ID) Then
        ObtenerAcoplamientoDirecto = m_AcoplamientosDirecto(m_AcoplamientosDirectoIdx(ID))
    Else
        Err.Raise 5, , "Acoplamiento Directo ID " & ID & " no encontrado"
    End If
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "clsCatalogos.ObtenerAcoplamientoDirecto", Err.Description
End Function

Public Function ObtenerReductor(ByVal ID As Long) As tCatalogReductores
    On Error GoTo ErrHandler
    If m_ReductoresIdx.Exists(ID) Then
        ObtenerReductor = m_Reductores(m_ReductoresIdx(ID))
    Else
        Err.Raise 5, , "Reductor ID " & ID & " no encontrado"
    End If
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "clsCatalogos.ObtenerReductor", Err.Description
End Function

Public Function ObtenerValvulaRetencion(ByVal ID As Long) As tCatalogValvulasRetencion
    On Error GoTo ErrHandler
    If m_ValvulasRetencionIdx.Exists(ID) Then
        ObtenerValvulaRetencion = m_ValvulasRetencion(m_ValvulasRetencionIdx(ID))
    Else
        Err.Raise 5, , "Válvula Retención ID " & ID & " no encontrada"
    End If
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "clsCatalogos.ObtenerValvulaRetencion", Err.Description
End Function

Public Function ObtenerElectrovalvula(ByVal ID As Long) As tCatalogElectrovalvulas
    On Error GoTo ErrHandler
    If m_ElectrovalvulasIdx.Exists(ID) Then
        ObtenerElectrovalvula = m_Electrovalvulas(m_ElectrovalvulasIdx(ID))
    Else
        Err.Raise 5, , "Electroválvula ID " & ID & " no encontrada"
    End If
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "clsCatalogos.ObtenerElectrovalvula", Err.Description
End Function

Public Function ObtenerFiltro(ByVal ID As Long) As tCatalogFiltros
    On Error GoTo ErrHandler
    If m_FiltrosIdx.Exists(ID) Then
        ObtenerFiltro = m_Filtros(m_FiltrosIdx(ID))
    Else
        Err.Raise 5, , "Filtro ID " & ID & " no encontrado"
    End If
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "clsCatalogos.ObtenerFiltro", Err.Description
End Function

Public Function ObtenerManometro(ByVal ID As Long) As tCatalogManometros
    On Error GoTo ErrHandler
    If m_ManometrosIdx.Exists(ID) Then
        ObtenerManometro = m_Manometros(m_ManometrosIdx(ID))
    Else
        Err.Raise 5, , "Manómetro ID " & ID & " no encontrado"
    End If
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "clsCatalogos.ObtenerManometro", Err.Description
End Function

'=========================================================
' FUNCIONES AUXILIARES
'=========================================================

Private Function NzVBA(ByVal v As Variant) As String
    If IsNull(v) Then NzVBA = "" Else NzVBA = CStr(v)
End Function

Private Function NzLong(ByVal v As Variant) As Long
    If IsNull(v) Then NzLong = 0 Else NzLong = CLng(v)
End Function

Private Function NzDouble(ByVal v As Variant) As Double
    If IsNull(v) Then NzDouble = 0 Else NzDouble = CDbl(v)
End Function

