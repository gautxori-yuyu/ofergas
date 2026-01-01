Attribute VB_Name = "modTestEjemplo"
' modTestEjemplo.bas - Ejemplo fiel a tu Access
Option Explicit
Sub Test_CrearOfertaCompleta()
    Dim conn As clsDBManager: Set conn = New clsDBManager
    conn.Connect ThisWorkbook.Path & "\Ofertas_Ejemplo.accdb"
    
    Dim dao As clsGenericDAO: Set dao = New clsGenericDAO
    Set dao.DBConnection = conn
    
    ' === DATOS COMPLETOS (estructura anidada) ===
    Dim maestra As Object
    Set maestra = CreateObject("Scripting.Dictionary")
    maestra("OFER_NUM_OFERTA") = "OF-TEST-100"
    maestra("OFER_FECHA") = Date
    maestra("OFER_CLIENTE") = "Cliente de Prueba"
    maestra("GASE_ID") = 1
    maestra("PAIS_ID") = 1
    
    ' OfertasManoObra (1:N)
    Dim manoObra As Object
    Set manoObra = CreateObject("System.Collections.ArrayList")
    Dim linea1 As Object
    Set linea1 = CreateObject("Scripting.Dictionary")
    linea1("OFMO_DESCRIPCION") = "Instalación eléctrica"
    linea1("OFMO_HORAS") = 8
    manoObra.Add linea1
    
    Dim linea2 As Object
    Set linea2 = CreateObject("Scripting.Dictionary")
    linea2("OFMO_DESCRIPCION") = "Puesta en marcha"
    linea2("OFMO_HORAS") = 4
    manoObra.Add linea2
    
    ' === INSERTAR ===
    On Error GoTo ErrorHandler
    
    ' 1. Insertar maestra ? obtener OFER_ID
    Dim resultMaestra As Object
    Set resultMaestra = dao.InsertRecord("OfertasDatosGenerales", maestra)
    If resultMaestra Is Nothing Then
        Err.Raise vbObjectError + 5200, , "Fallo al insertar maestra"
    End If
    
    Dim idOferta As String
    If resultMaestra.Exists("OFER_ID") Then
        idOferta = resultMaestra("OFER_ID")
    Else
        Err.Raise vbObjectError + 5201, , "No se recuperó OFER_ID"
    End If
    
    ' 2. Propagar OFER_ID a hijas y insertar
    Dim i As Long
    For i = 0 To manoObra.Count - 1
        Dim reg As Object
        Set reg = manoObra(i)
        If Not reg.Exists("OFER_ID") Then
            reg.Add "OFER_ID", idOferta
        End If
    Next i
    
    Dim resultManoObra As Object
    Set resultManoObra = dao.InsertRecord("OfertasManoObra", manoObra)
    If resultManoObra Is Nothing Then
        Err.Raise vbObjectError + 5202, , "Fallo al insertar OfertasManoObra"
    End If
    
    ' 3. Mostrar resultados
    Debug.Print "[OK] OFER_ID generado: " & idOferta
    If resultManoObra.Exists("OFMO_ID") Then
        Debug.Print "[OK] Primer OFMO_ID: " & resultManoObra("OFMO_ID")
    End If
    
    conn.Disconnect
    MsgBox "? Oferta completa creada. Ver Debug.Print.", vbInformation
    Exit Sub
    
ErrorHandler:
    If Not conn Is Nothing Then
        On Error Resume Next
        conn.Disconnect
    End If
    MsgBox "Error: " & Err.Description & " (Err=" & Err.Number & ")", vbCritical
End Sub

Sub DuplicarOfertaExistente()
    Dim conn As clsDBManager: Set conn = New clsDBManager
    conn.Connect ThisWorkbook.Path & "\Ofertas_Ejemplo.accdb"
    
    Dim dao As clsGenericDAO: Set dao = New clsGenericDAO
    Set dao.DBConnection = conn
    
    Dim oferta As clsOferta: Set oferta = New clsOferta
    Set oferta.GenericDAO = dao
    
    Dim idOrigen As String
    idOrigen = "{11111111-1111-1111-1111-111111111111}"
    
    Dim idNuevo As String
    idNuevo = oferta.DuplicarOferta(idOrigen, "OF-COPIA-001")
    
    conn.Disconnect
End Sub
Private Function GetChildTables() As Variant
    GetChildTables = Array("OfertasManoObra", "OfertasModelo", "OfertasOpciones", _
        "OfertasOtros", "OfertasRefrigeradores", "OfertasAccesorios", _
        "OfertasCabezal", "OfertasCalderines", "OfertasExtras", _
        "OfertasInstrumentacion")
End Function

' ══════════════════════════════════════════════════════════
' EJEMPLOS DE USO DE FUNCIONALIDADES AVANZADAS
' ══════════════════════════════════════════════════════════

Sub Test_GetRecord()
    ' Obtener UN registro específico por ID
    Dim conn As clsDBManager: Set conn = New clsDBManager
    conn.Connect ThisWorkbook.Path & "\Ofertas_Ejemplo.accdb"

    Dim dao As clsGenericDAO: Set dao = New clsGenericDAO
    Set dao.DBConnection = conn

    Dim oferta As Object
    Set oferta = dao.GetRecord("OfertasDatosGenerales", "{11111111-1111-1111-1111-111111111111}", "OFER_ID")

    If Not oferta Is Nothing Then
        Debug.Print "[OK] Oferta encontrada: " & oferta("OFER_NUM_OFERTA")
        Debug.Print "    Cliente: " & oferta("OFER_CLIENTE")
    Else
        Debug.Print "[WARN] Oferta no encontrada"
    End If

    conn.Disconnect
End Sub

Sub Test_SelectRecords()
    ' Obtener múltiples registros con filtro, orden y límite
    Dim conn As clsDBManager: Set conn = New clsDBManager
    conn.Connect ThisWorkbook.Path & "\Ofertas_Ejemplo.accdb"

    Dim dao As clsGenericDAO: Set dao = New clsGenericDAO
    Set dao.DBConnection = conn

    ' Obtener últimas 10 ofertas de 2025
    Dim ofertas As Object
    Set ofertas = dao.SelectRecords("OfertasDatosGenerales", _
                                    "YEAR(OFER_FECHA) = 2025", _
                                    "OFER_FECHA DESC", _
                                    10)

    Debug.Print "[OK] Encontradas " & ofertas.Count & " ofertas"
    Dim i As Long
    For i = 0 To ofertas.Count - 1
        Debug.Print "  " & ofertas(i)("OFER_NUM_OFERTA") & " - " & ofertas(i)("OFER_CLIENTE")
    Next i

    conn.Disconnect
End Sub

Sub Test_SearchRecords()
    ' Búsqueda avanzada con operadores
    Dim conn As clsDBManager: Set conn = New clsDBManager
    conn.Connect ThisWorkbook.Path & "\Ofertas_Ejemplo.accdb"

    Dim dao As clsGenericDAO: Set dao = New clsGenericDAO
    Set dao.DBConnection = conn

    ' Crear filtros
    Dim filters As Object
    Set filters = CreateObject("Scripting.Dictionary")

    ' OFER_FECHA > 2025-01-01
    Dim filtroFecha As Object
    Set filtroFecha = CreateObject("Scripting.Dictionary")
    filtroFecha.Add ">", #1/1/2025#
    filters.Add "OFER_FECHA", filtroFecha

    ' OFER_CLIENTE LIKE '%Acme%'
    Dim filtroCliente As Object
    Set filtroCliente = CreateObject("Scripting.Dictionary")
    filtroCliente.Add "LIKE", "%Acme%"
    filters.Add "OFER_CLIENTE", filtroCliente

    Dim ofertas As Object
    Set ofertas = dao.SearchRecords("OfertasDatosGenerales", filters, "OFER_FECHA DESC", 20)

    Debug.Print "[OK] Búsqueda avanzada: " & ofertas.Count & " resultados"

    conn.Disconnect
End Sub

Sub Test_UpdateRecord()
    ' Actualizar registro(s)
    Dim conn As clsDBManager: Set conn = New clsDBManager
    conn.Connect ThisWorkbook.Path & "\Ofertas_Ejemplo.accdb"

    Dim dao As clsGenericDAO: Set dao = New clsGenericDAO
    Set dao.DBConnection = conn

    ' Datos a actualizar
    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")
    data.Add "OFER_CLIENTE", "Cliente Actualizado S.L."
    data.Add "OFER_OBSERVACIONES", "Actualizado automáticamente"

    ' WHERE clause
    Dim where As String
    where = "[OFER_ID] = '{11111111-1111-1111-1111-111111111111}'"

    If dao.UpdateRecord("OfertasDatosGenerales", data, where) Then
        Debug.Print "[OK] Registro actualizado"
    Else
        Debug.Print "[ERR] Fallo al actualizar"
    End If

    conn.Disconnect
End Sub

Sub Test_DeleteRecordCascade()
    ' Eliminar oferta con todas sus tablas hijas
    Dim conn As clsDBManager: Set conn = New clsDBManager
    conn.Connect ThisWorkbook.Path & "\Ofertas_Ejemplo.accdb"

    Dim dao As clsGenericDAO: Set dao = New clsGenericDAO
    Set dao.DBConnection = conn

    ' Definir tablas hijas
    Dim childTables As Variant
    childTables = GetChildTablesArray()

    ' Eliminar con cascade
    Dim idToDelete As String
    idToDelete = "{99999999-9999-9999-9999-999999999999}"

    If dao.DeleteRecordCascade("OfertasDatosGenerales", idToDelete, "OFER_ID", childTables) Then
        Debug.Print "[OK] Oferta eliminada con todas sus hijas"
    Else
        Debug.Print "[ERR] Fallo al eliminar"
    End If

    conn.Disconnect
End Sub

Sub Test_BulkInsert()
    ' Inserción masiva optimizada
    Dim conn As clsDBManager: Set conn = New clsDBManager
    conn.Connect ThisWorkbook.Path & "\Ofertas_Ejemplo.accdb"

    Dim dao As clsGenericDAO: Set dao = New clsGenericDAO
    Set dao.DBConnection = conn

    ' Crear 100 registros de mano de obra
    Dim data As Object
    Set data = CreateObject("System.Collections.ArrayList")

    Dim i As Long
    For i = 1 To 100
        Dim reg As Object
        Set reg = CreateObject("Scripting.Dictionary")
        reg.Add "OFER_ID", "{11111111-1111-1111-1111-111111111111}"
        reg.Add "OFMO_DESCRIPCION", "Trabajo " & i
        reg.Add "OFMO_HORAS", i Mod 10 + 1
        data.Add reg
    Next i

    If dao.BulkInsert("OfertasManoObra", data) Then
        Debug.Print "[OK] " & data.Count & " registros insertados masivamente"
    Else
        Debug.Print "[ERR] Fallo en BulkInsert"
    End If

    conn.Disconnect
End Sub

Sub Test_ValidateForeignKeys()
    ' Validar FKs antes de insertar
    Dim conn As clsDBManager: Set conn = New clsDBManager
    conn.Connect ThisWorkbook.Path & "\Ofertas_Ejemplo.accdb"

    Dim dao As clsGenericDAO: Set dao = New clsGenericDAO
    Set dao.DBConnection = conn

    ' Datos a insertar
    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")
    data.Add "OFER_NUM_OFERTA", "OF-TEST-999"
    data.Add "OFER_FECHA", Date
    data.Add "GASE_ID", 99999  ' FK inválida
    data.Add "PAIS_ID", 1

    ' Definir validaciones FK
    Dim fkValidations(0 To 1) As Object
    Set fkValidations(0) = CreateObject("Scripting.Dictionary")
    fkValidations(0).Add "campo", "GASE_ID"
    fkValidations(0).Add "tablaRef", "GamasServicios"
    fkValidations(0).Add "campoRef", "gase_id"

    Set fkValidations(1) = CreateObject("Scripting.Dictionary")
    fkValidations(1).Add "campo", "PAIS_ID"
    fkValidations(1).Add "tablaRef", "Paises"
    fkValidations(1).Add "campoRef", "pais_id"

    ' Validar
    Dim errors As Object
    Set errors = dao.ValidateForeignKeys(data, fkValidations)

    If errors.Count > 0 Then
        Debug.Print "[ERR] Se encontraron " & errors.Count & " errores de FK:"
        Dim campo As Variant
        For Each campo In errors.Keys
            Debug.Print "  - " & errors(campo)
        Next campo
    Else
        Debug.Print "[OK] Todas las FKs son válidas"
    End If

    conn.Disconnect
End Sub

Sub Test_ExportImportExcel()
    ' Exportar e importar desde Excel
    Dim conn As clsDBManager: Set conn = New clsDBManager
    conn.Connect ThisWorkbook.Path & "\Ofertas_Ejemplo.accdb"

    Dim dao As clsGenericDAO: Set dao = New clsGenericDAO
    Set dao.DBConnection = conn

    ' 1. Obtener datos
    Dim ofertas As Object
    Set ofertas = dao.SelectRecords("OfertasDatosGenerales", "", "OFER_FECHA DESC", 50)

    ' 2. Exportar a Excel
    Dim excelPath As String
    excelPath = ThisWorkbook.Path & "\Ofertas_Export.xlsx"

    If ExportToExcel("OfertasDatosGenerales", ofertas, excelPath) Then
        Debug.Print "[OK] Exportado a " & excelPath

        ' 3. Importar desde Excel
        Dim importData As Object
        Set importData = ImportFromExcel(excelPath)

        If Not importData Is Nothing Then
            Debug.Print "[OK] Importados " & importData.Count & " registros desde Excel"
        End If
    End If

    conn.Disconnect
End Sub

Sub Test_ExportImportCSV()
    ' Exportar e importar desde CSV
    Dim conn As clsDBManager: Set conn = New clsDBManager
    conn.Connect ThisWorkbook.Path & "\Ofertas_Ejemplo.accdb"

    Dim dao As clsGenericDAO: Set dao = New clsGenericDAO
    Set dao.DBConnection = conn

    ' 1. Obtener datos
    Dim manoObra As Object
    Set manoObra = dao.SelectRecords("OfertasManoObra", "", "OFMO_ID", 100)

    ' 2. Exportar a CSV
    Dim csvPath As String
    csvPath = ThisWorkbook.Path & "\ManoObra_Export.csv"

    If ExportToCSV("OfertasManoObra", manoObra, csvPath) Then
        Debug.Print "[OK] Exportado a " & csvPath

        ' 3. Importar desde CSV
        Dim importData As Object
        Set importData = ImportFromCSV(csvPath)

        If Not importData Is Nothing Then
            Debug.Print "[OK] Importados " & importData.Count & " registros desde CSV"
        End If
    End If

    conn.Disconnect
End Sub

Sub Test_BackupRestore()
    ' Backup y restore de base de datos usando clsDBManager
    Dim dbPath As String
    dbPath = ThisWorkbook.Path & "\Ofertas_Ejemplo.accdb"

    Dim backupPath As String
    backupPath = ThisWorkbook.Path & "\Backups\Ofertas_" & Format(Now, "yyyymmdd_hhnnss") & ".accdb"

    ' Crear instancia de DBManager
    Dim dbMgr As clsDBManager
    Set dbMgr = New clsDBManager
    dbMgr.Connect dbPath

    ' 1. Crear backup
    If dbMgr.Backup(backupPath) Then
        Debug.Print "[OK] Backup creado en " & backupPath

        ' 2. Restaurar (ejemplo)
        ' Dim restorePath As String
        ' restorePath = ThisWorkbook.Path & "\Ofertas_Restored.accdb"
        ' If dbMgr.Restore(backupPath, restorePath) Then
        '     Debug.Print "[OK] Base de datos restaurada"
        ' End If
    End If

    dbMgr.Disconnect
End Sub

Private Function GetChildTablesArray() As Variant
    ' Devuelve array de Dictionary para cascade delete
    Dim tables(0 To 9) As Object

    Set tables(0) = CreateObject("Scripting.Dictionary")
    tables(0).Add "tabla", "OfertasManoObra"
    tables(0).Add "fkField", "OFER_ID"

    Set tables(1) = CreateObject("Scripting.Dictionary")
    tables(1).Add "tabla", "OfertasModelo"
    tables(1).Add "fkField", "OFER_ID"

    Set tables(2) = CreateObject("Scripting.Dictionary")
    tables(2).Add "tabla", "OfertasOpciones"
    tables(2).Add "fkField", "OFER_ID"

    Set tables(3) = CreateObject("Scripting.Dictionary")
    tables(3).Add "tabla", "OfertasOtros"
    tables(3).Add "fkField", "OFER_ID"

    Set tables(4) = CreateObject("Scripting.Dictionary")
    tables(4).Add "tabla", "OfertasRefrigeradores"
    tables(4).Add "fkField", "OFER_ID"

    Set tables(5) = CreateObject("Scripting.Dictionary")
    tables(5).Add "tabla", "OfertasAccesorios"
    tables(5).Add "fkField", "OFER_ID"

    Set tables(6) = CreateObject("Scripting.Dictionary")
    tables(6).Add "tabla", "OfertasCabezal"
    tables(6).Add "fkField", "OFER_ID"

    Set tables(7) = CreateObject("Scripting.Dictionary")
    tables(7).Add "tabla", "OfertasCalderines"
    tables(7).Add "fkField", "OFER_ID"

    Set tables(8) = CreateObject("Scripting.Dictionary")
    tables(8).Add "tabla", "OfertasExtras"
    tables(8).Add "fkField", "OFER_ID"

    Set tables(9) = CreateObject("Scripting.Dictionary")
    tables(9).Add "tabla", "OfertasInstrumentacion"
    tables(9).Add "fkField", "OFER_ID"

    GetChildTablesArray = tables
End Function
