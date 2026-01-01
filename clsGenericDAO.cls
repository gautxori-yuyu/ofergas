VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGenericDAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_dbConn As clsDBManager

Public Property Set DBConnection(ByVal conn As clsDBManager)
    Set m_dbConn = conn
End Property

Public Property Get DBConnection() As clsDBManager
    Set DBConnection = m_dbConn
End Property
' ══════════════════════════════════════════════════════════
' INSERTA un registro (1:1) o múltiples (1:N)
' data:
'   - Dictionary ? 1 registro
'   - System.Collections.ArrayList de Dictionary ? N registros
' Devuelve un Dictionary con {campoAutonum => valorGenerado} del primer registro,
' o Nothing si fallo. Para múltiples registros, solo devuelve del primero (como Access).
' ══════════════════════════════════════════════════════════
Public Function InsertRecord(ByVal tableName As String, ByVal data As Object, _
                              Optional ByVal idFieldName As String = "OFER_ID") As Object
    On Error GoTo ErrorHandler
    
    Dim autoFields As Variant
    autoFields = GetAutoNumericFields(tableName)

    Dim camposDestino As String
    camposDestino = GetFieldsExcept(tableName, autoFields)
    If camposDestino = "" Then
        Set InsertRecord = Nothing
        Exit Function
    End If
    
    Dim esLista As Boolean
    esLista = TypeName(data) = "ArrayList"
    
    Dim listaRegistros As Object
    If esLista Then
        Set listaRegistros = data
    Else
        Set listaRegistros = CreateObject("System.Collections.ArrayList")
        listaRegistros.Add data
    End If
    
    Dim results As Object
    Set results = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 0 To listaRegistros.Count - 1
        Dim reg As Object: Set reg = listaRegistros(i)
        Dim camposArray() As String
        camposArray = Split(camposDestino, ",")
        
        Dim listaCampos As String, listaValores As String
        Dim j As Long
        For j = 0 To UBound(camposArray)
            Dim campoBD As String
            campoBD = Trim$(Replace$(Replace$(camposArray(j), "[", ""), "]", ""))
            Dim valorSQL As String
            
            If reg.Exists(campoBD) Then
                valorSQL = EscapeValue(reg(campoBD))
            Else
                valorSQL = "NULL"
            End If
            
            If listaCampos <> "" Then listaCampos = listaCampos & ","
            listaCampos = listaCampos & "[" & campoBD & "]"
            
            If listaValores <> "" Then listaValores = listaValores & ","
            listaValores = listaValores & valorSQL
        Next j
        
        Dim sql As String
        sql = "INSERT INTO [" & tableName & "] (" & listaCampos & ") VALUES (" & listaValores & ")"
        m_dbConn.Connection.Execute sql
        
        ' Solo recuperar autonuméricos del PRIMER registro (como hace tu Access)
        If i = 0 And UBound(autoFields) >= LBound(autoFields) Then
            Dim k As Long
            For k = LBound(autoFields) To UBound(autoFields)
                Dim campoAuto As String: campoAuto = autoFields(k)
                Dim valorGenerado As String
                ' Usar GetLastAutoID con los campos del registro como uniqueFields
                valorGenerado = GetLastAutoID(tableName, campoAuto, reg)
                If valorGenerado <> "" Then
                    results.Add campoAuto, valorGenerado
                End If
            Next k
        End If
    Next i
    
    Set InsertRecord = results
    Exit Function
    
ErrorHandler:
    Debug.Print "[ERR] InsertRecord(" & tableName & "): " & Err.Description
    Set InsertRecord = Nothing
End Function

' ══════════════════════════════════════════════════════════
' DUPLICAR genérico (como tu DuplicarRegistroEnTablas)
' ══════════════════════════════════════════════════════════
Public Function DuplicateRecord(ByVal tableName As String, ByVal idOrigen As String, ByVal idNuevo As String, ByVal idFieldName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim autoFields As Variant
    autoFields = GetAutoNumericFields(tableName)

    Dim camposDestino As String
    camposDestino = GetFieldsExcept(tableName, autoFields)
    If camposDestino = "" Then
        DuplicateRecord = True
        Exit Function
    End If
    
    Dim camposOrigen As String
    camposOrigen = camposDestino
    camposOrigen = ReplaceFieldInSQL(camposOrigen, idFieldName, "'" & idNuevo & "'")
    
    Dim sql As String
    sql = "INSERT INTO [" & tableName & "] (" & camposDestino & ") " & _
          "SELECT " & camposOrigen & " FROM [" & tableName & "] WHERE [" & idFieldName & "] = '" & EscapeString(idOrigen) & "'"
    
    m_dbConn.Connection.Execute sql
    DuplicateRecord = True
    Exit Function
    
ErrorHandler:
    Debug.Print "[ERR] DuplicateRecord(" & tableName & "): " & Err.Description
    DuplicateRecord = False
End Function

' ══════════════════════════════════════════════════════════
' DUPLICAR AVANZADO - Con WHERE dinámico y nuevos valores
' ══════════════════════════════════════════════════════════
Public Function DuplicateRecordAdvanced(ByVal tableName As String, _
                                       ByVal whereClause As String, _
                                       ByVal newValues As Object, _
                                       Optional ByVal excludeFields As Variant) As Boolean
    ' whereClause: Ej: "OFER_NUM_OFERTA = 'OF-001'"
    ' newValues: Dictionary {"OFER_NUM_OFERTA": "OF-002", "OFER_FECHA": Date}
    ' excludeFields: Array de campos a excluir (además de autonuméricos)
    On Error GoTo ErrorHandler

    ' Obtener campos autonuméricos a excluir
    Dim autoFields As Variant
    autoFields = GetAutoNumericFields(tableName)

    ' Combinar autoFields con excludeFields
    Dim allExclude As Variant
    If Not IsMissing(excludeFields) Then
        If IsArray(excludeFields) Then
            ' Combinar arrays
            allExclude = CombineArrays(autoFields, excludeFields)
        Else
            ' excludeFields es un solo valor
            ReDim allExclude(UBound(autoFields) + 1)
            Dim i As Long
            For i = 0 To UBound(autoFields)
                allExclude(i) = autoFields(i)
            Next i
            allExclude(UBound(allExclude)) = excludeFields
        End If
    Else
        allExclude = autoFields
    End If

    ' Obtener campos de destino (excluyendo autonuméricos y excludeFields)
    Dim camposDestino As String
    camposDestino = GetFieldsExcept(tableName, allExclude)
    If camposDestino = "" Then
        DuplicateRecordAdvanced = True
        Exit Function
    End If

    ' Construir SELECT con nuevos valores
    Dim camposOrigen As String
    camposOrigen = camposDestino

    ' Reemplazar campos con nuevos valores
    If Not newValues Is Nothing Then
        Dim campo As Variant
        For Each campo In newValues.Keys
            camposOrigen = ReplaceFieldInSQL(camposOrigen, CStr(campo), EscapeValue(newValues(campo)))
        Next campo
    End If

    ' Construir y ejecutar SQL
    Dim sql As String
    sql = "INSERT INTO [" & tableName & "] (" & camposDestino & ") " & _
          "SELECT " & camposOrigen & " FROM [" & tableName & "] WHERE " & whereClause

    m_dbConn.Connection.Execute sql
    DuplicateRecordAdvanced = True
    Debug.Print "[OK] DuplicateRecordAdvanced: " & tableName
    Exit Function

ErrorHandler:
    Debug.Print "[ERR] DuplicateRecordAdvanced(" & tableName & "): " & Err.Description
    DuplicateRecordAdvanced = False
End Function

' ══════════════════════════════════════════════════════════
' BORRAR genérico
' ══════════════════════════════════════════════════════════
Public Function DeleteRecord(ByVal tableName As String, ByVal idValue As String, ByVal idFieldName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim sql As String
    sql = "DELETE FROM [" & tableName & "] WHERE [" & idFieldName & "] = '" & EscapeString(idValue) & "'"
    m_dbConn.Connection.Execute sql
    DeleteRecord = True
    Exit Function

ErrorHandler:
    Debug.Print "[ERR] DeleteRecord(" & tableName & "): " & Err.Description
    DeleteRecord = False
End Function

' ══════════════════════════════════════════════════════════
' RECUPERAR valor de un campo autonumérico específico (no solo OFER_ID)
' ══════════════════════════════════════════════════════════
' En clsGenericDAO.cls
Public Function GetLastAutoID(ByVal tableName As String, ByVal idFieldName As String, ByVal uniqueFields As Object) As String
    ' uniqueFields: Dictionary {"campo1": valor1, "campo2": valor2}
    Dim sql As String
    sql = "SELECT TOP 1 [" & idFieldName & "] FROM [" & tableName & "] WHERE "
    
    Dim campo As Variant, condiciones As String
    For Each campo In uniqueFields.Keys
        Dim cond As String
        cond = "[" & campo & "] = " & EscapeValue(uniqueFields(campo))
        If condiciones <> "" Then condiciones = condiciones & " AND "
        condiciones = condiciones & cond
    Next campo
    
    sql = sql & condiciones & " ORDER BY [" & idFieldName & "] DESC"
    
    Dim rs As Object
    Set rs = m_dbConn.Connection.Execute(sql)
    If Not rs.EOF Then
        GetLastAutoID = NzVBA(rs(0), "")
        ' Limpiar formato Access: {guid {...}} ? {...}
        If InStr(GetLastAutoID, "{guid {") > 0 Then
            GetLastAutoID = Replace(Replace(GetLastAutoID, "{guid {", "{"), "}}", "}")
        End If
    End If
    rs.Close
End Function

' ══════════════════════════════════════════════════════════
' FUNCIONES AUXILIARES
' ══════════════════════════════════════════════════════════

Private Function EscapeValue(ByVal val As Variant) As String
    If IsNull(val) Or val = "" Then
        EscapeValue = "NULL"
    ElseIf VarType(val) = vbString Then
        EscapeValue = "'" & Replace(val, "'", "''") & "'"
    ElseIf VarType(val) = vbDate Then
        EscapeValue = "#" & Format(val, "yyyy\/mm\/dd hh\:nn\:ss") & "#"
    ElseIf VarType(val) = vbBoolean Then
        EscapeValue = IIf(val, "True", "False")
    Else
        EscapeValue = CStr(val)
    End If
End Function

Private Function EscapeString(ByVal s As String) As String
    EscapeString = "'" & Replace(s, "'", "''") & "'"
End Function

Private Function EscapeDate(ByVal d As Date) As String
    EscapeDate = "#" & Format(d, "yyyy\/mm\/dd hh\:nn\:ss") & "#"
End Function

Private Function IsInArray(ByVal value As String, ByRef arr As Variant) As Boolean
    If Not IsArray(arr) Then Exit Function
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If StrComp(value, arr(i), vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next i
End Function

' ══════════════════════════════════════════════════════════
' SELECT/READ GENÉRICO - Recuperar registros
' ══════════════════════════════════════════════════════════
Public Function SelectRecords(ByVal tableName As String, _
                              Optional ByVal whereClause As String = "", _
                              Optional ByVal orderBy As String = "", _
                              Optional ByVal limit As Long = 0) As Object
    On Error GoTo ErrorHandler

    Dim sql As String
    sql = "SELECT "

    ' Aplicar límite si se especifica (TOP N en Access)
    If limit > 0 Then sql = sql & "TOP " & limit & " "

    sql = sql & "* FROM [" & tableName & "]"

    ' Aplicar filtro WHERE si existe
    If whereClause <> "" Then
        sql = sql & " WHERE " & whereClause
    End If

    ' Aplicar ordenamiento si existe
    If orderBy <> "" Then
        sql = sql & " ORDER BY " & orderBy
    End If

    Dim rs As Object
    Set rs = m_dbConn.Connection.Execute(sql)

    ' Convertir recordset a ArrayList de Dictionary
    Dim results As Object
    Set results = CreateObject("System.Collections.ArrayList")

    Do While Not rs.EOF
        Dim record As Object
        Set record = CreateObject("Scripting.Dictionary")

        Dim i As Long
        For i = 0 To rs.Fields.Count - 1
            record.Add rs.Fields(i).Name, NzVariant(rs.Fields(i).Value)
        Next i

        results.Add record
        rs.MoveNext
    Loop

    rs.Close
    Set SelectRecords = results
    Exit Function

ErrorHandler:
    Debug.Print "[ERR] SelectRecords(" & tableName & "): " & Err.Description
    Set SelectRecords = CreateObject("System.Collections.ArrayList")
End Function

' ══════════════════════════════════════════════════════════
' GET RECORD - Obtener un solo registro por ID
' ══════════════════════════════════════════════════════════
Public Function GetRecord(ByVal tableName As String, _
                         ByVal idValue As String, _
                         ByVal idFieldName As String) As Object
    On Error GoTo ErrorHandler

    Dim whereClause As String
    whereClause = "[" & idFieldName & "] = " & EscapeValue(idValue)

    Dim results As Object
    Set results = SelectRecords(tableName, whereClause, "", 1)

    If results.Count > 0 Then
        Set GetRecord = results(0)
    Else
        Set GetRecord = Nothing
    End If
    Exit Function

ErrorHandler:
    Debug.Print "[ERR] GetRecord(" & tableName & "): " & Err.Description
    Set GetRecord = Nothing
End Function

' ══════════════════════════════════════════════════════════
' BÚSQUEDA/FILTRADO AVANZADO con operadores dinámicos
' filters: Dictionary con {campo: {operador: valor}}
' Ejemplo: {"OFER_FECHA": {">": #2025/01/01#}, "OFER_CLIENTE": {"LIKE": "%Acme%"}}
' ══════════════════════════════════════════════════════════
Public Function SearchRecords(ByVal tableName As String, _
                              ByVal filters As Object, _
                              Optional ByVal orderBy As String = "", _
                              Optional ByVal limit As Long = 0) As Object
    On Error GoTo ErrorHandler

    ' Construir cláusula WHERE dinámica
    Dim whereClause As String
    Dim campo As Variant

    For Each campo In filters.Keys
        Dim filterDef As Object
        Set filterDef = filters(campo)

        Dim operador As Variant
        For Each operador In filterDef.Keys
            Dim valor As Variant
            valor = filterDef(operador)

            If whereClause <> "" Then whereClause = whereClause & " AND "

            ' Construir condición según operador
            Select Case UCase(Trim$(operador))
                Case "=", "<", ">", "<=", ">=", "<>"
                    whereClause = whereClause & "[" & campo & "] " & operador & " " & EscapeValue(valor)
                Case "LIKE"
                    whereClause = whereClause & "[" & campo & "] LIKE " & EscapeValue(valor)
                Case "IN"
                    ' valor debe ser Array
                    whereClause = whereClause & "[" & campo & "] IN (" & BuildINList(valor) & ")"
                Case "BETWEEN"
                    ' valor debe ser Array(min, max)
                    whereClause = whereClause & "[" & campo & "] BETWEEN " & EscapeValue(valor(0)) & " AND " & EscapeValue(valor(1))
                Case "IS NULL"
                    whereClause = whereClause & "[" & campo & "] IS NULL"
                Case "IS NOT NULL"
                    whereClause = whereClause & "[" & campo & "] IS NOT NULL"
            End Select
        Next operador
    Next campo

    ' Usar SelectRecords con WHERE construido
    Set SearchRecords = SelectRecords(tableName, whereClause, orderBy, limit)
    Exit Function

ErrorHandler:
    Debug.Print "[ERR] SearchRecords(" & tableName & "): " & Err.Description
    Set SearchRecords = CreateObject("System.Collections.ArrayList")
End Function

' ══════════════════════════════════════════════════════════
' UPDATE GENÉRICO - Actualizar registro(s)
' ══════════════════════════════════════════════════════════
Public Function UpdateRecord(ByVal tableName As String, _
                            ByVal data As Object, _
                            ByVal whereClause As String) As Boolean
    On Error GoTo ErrorHandler

    ' Construir SET clause
    Dim setClause As String
    Dim campo As Variant

    For Each campo In data.Keys
        If setClause <> "" Then setClause = setClause & ", "
        setClause = setClause & "[" & campo & "] = " & EscapeValue(data(campo))
    Next campo

    If setClause = "" Then
        UpdateRecord = False
        Exit Function
    End If

    ' Construir y ejecutar SQL
    Dim sql As String
    sql = "UPDATE [" & tableName & "] SET " & setClause

    If whereClause <> "" Then
        sql = sql & " WHERE " & whereClause
    End If

    m_dbConn.Connection.Execute sql
    UpdateRecord = True
    Exit Function

ErrorHandler:
    Debug.Print "[ERR] UpdateRecord(" & tableName & "): " & Err.Description
    UpdateRecord = False
End Function

Private Function NzVariant(ByVal v As Variant) As Variant
    If IsNull(v) Then
        NzVariant = Empty
    Else
        NzVariant = v
    End If
End Function

Private Function BuildINList(ByVal arr As Variant) As String
    If Not IsArray(arr) Then Exit Function

    Dim i As Long, lista As String
    For i = LBound(arr) To UBound(arr)
        If lista <> "" Then lista = lista & ","
        lista = lista & EscapeValue(arr(i))
    Next i

    BuildINList = lista
End Function

' ══════════════════════════════════════════════════════════
' CASCADE DELETE - Eliminar registro principal y todos sus hijos
' childTables: Array de Dictionary con {tabla: "NombreTabla", fkField: "CAMPO_FK"}
' Ejemplo: Array(Dict("tabla"="OfertasManoObra", "fkField"="OFER_ID"))
' ══════════════════════════════════════════════════════════
Public Function DeleteRecordCascade(ByVal tableName As String, _
                                   ByVal idValue As String, _
                                   ByVal idFieldName As String, _
                                   ByVal childTables As Variant) As Boolean
    On Error GoTo ErrorHandler

    ' Iniciar transacción
    Dim ws As Object
    Set ws = m_dbConn.DAOWorkspace
    ws.BeginTrans

    ' 1. Eliminar todos los registros hijos
    If IsArray(childTables) Then
        Dim i As Long
        For i = LBound(childTables) To UBound(childTables)
            Dim childDef As Object
            Set childDef = childTables(i)

            Dim childTable As String, fkField As String
            childTable = childDef("tabla")
            fkField = childDef("fkField")

            ' Delete hijos
            Dim sqlChild As String
            sqlChild = "DELETE FROM [" & childTable & "] WHERE [" & fkField & "] = '" & EscapeString(idValue) & "'"
            m_dbConn.Connection.Execute sqlChild
            Debug.Print "[OK] Deleted children from " & childTable
        Next i
    End If

    ' 2. Eliminar registro principal
    Dim sqlMain As String
    sqlMain = "DELETE FROM [" & tableName & "] WHERE [" & idFieldName & "] = '" & EscapeString(idValue) & "'"
    m_dbConn.Connection.Execute sqlMain

    ws.CommitTrans
    DeleteRecordCascade = True
    Debug.Print "[OK] Cascade delete completed for " & tableName & " id=" & idValue
    Exit Function

ErrorHandler:
    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Rollback
    End If
    Debug.Print "[ERR] DeleteRecordCascade(" & tableName & "): " & Err.Description
    DeleteRecordCascade = False
End Function

' ══════════════════════════════════════════════════════════
' BULK INSERT - Inserción masiva optimizada con transacción
' data: ArrayList de Dictionary
' ══════════════════════════════════════════════════════════
Public Function BulkInsert(ByVal tableName As String, _
                          ByVal data As Object) As Boolean
    On Error GoTo ErrorHandler

    If TypeName(data) <> "ArrayList" Then
        Debug.Print "[ERR] BulkInsert: data debe ser ArrayList"
        BulkInsert = False
        Exit Function
    End If

    If data.Count = 0 Then
        BulkInsert = True
        Exit Function
    End If

    ' Iniciar transacción para performance
    Dim ws As Object
    Set ws = m_dbConn.DAOWorkspace
    ws.BeginTrans

    ' Obtener campos excluyendo autonuméricos
    Dim autoFields As Variant
    autoFields = GetAutoNumericFields(tableName)

    Dim camposDestino As String
    camposDestino = GetFieldsExcept(tableName, autoFields)

    If camposDestino = "" Then
        ws.Rollback
        BulkInsert = False
        Exit Function
    End If

    Dim camposArray() As String
    camposArray = Split(camposDestino, ",")

    ' Insertar cada registro
    Dim i As Long
    For i = 0 To data.Count - 1
        Dim reg As Object
        Set reg = data(i)

        Dim listaCampos As String, listaValores As String
        Dim j As Long

        For j = 0 To UBound(camposArray)
            Dim campoBD As String
            campoBD = Trim$(Replace$(Replace$(camposArray(j), "[", ""), "]", ""))

            Dim valorSQL As String
            If reg.Exists(campoBD) Then
                valorSQL = EscapeValue(reg(campoBD))
            Else
                valorSQL = "NULL"
            End If

            If listaCampos <> "" Then listaCampos = listaCampos & ","
            listaCampos = listaCampos & "[" & campoBD & "]"

            If listaValores <> "" Then listaValores = listaValores & ","
            listaValores = listaValores & valorSQL
        Next j

        Dim sql As String
        sql = "INSERT INTO [" & tableName & "] (" & listaCampos & ") VALUES (" & listaValores & ")"
        m_dbConn.Connection.Execute sql
    Next i

    ws.CommitTrans
    BulkInsert = True
    Debug.Print "[OK] BulkInsert: " & data.Count & " registros insertados en " & tableName
    Exit Function

ErrorHandler:
    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Rollback
    End If
    Debug.Print "[ERR] BulkInsert(" & tableName & "): " & Err.Description
    BulkInsert = False
End Function

' ══════════════════════════════════════════════════════════
' VALIDACIÓN DE FOREIGN KEYS - Verifica que FK exista en tabla referenciada
' fkValidations: Array de Dictionary con {campo, tablaRef, campoRef}
' Ejemplo: Array(Dict("campo"="PAIS_ID", "tablaRef"="Paises", "campoRef"="pais_id"))
' ══════════════════════════════════════════════════════════
Public Function ValidateForeignKeys(ByVal data As Object, _
                                    ByVal fkValidations As Variant) As Object
    On Error GoTo ErrorHandler

    ' Devuelve Dictionary con errores: {campo: mensajeError}
    Dim errors As Object
    Set errors = CreateObject("Scripting.Dictionary")

    If Not IsArray(fkValidations) Then
        Set ValidateForeignKeys = errors
        Exit Function
    End If

    Dim i As Long
    For i = LBound(fkValidations) To UBound(fkValidations)
        Dim fkDef As Object
        Set fkDef = fkValidations(i)

        Dim campo As String, tablaRef As String, campoRef As String
        campo = fkDef("campo")
        tablaRef = fkDef("tablaRef")
        campoRef = fkDef("campoRef")

        ' Si el campo existe en data y no es nulo/vacío
        If data.Exists(campo) Then
            Dim valorFK As Variant
            valorFK = data(campo)

            If Not IsNull(valorFK) And valorFK <> "" And valorFK <> 0 Then
                ' Verificar existencia en tabla referenciada
                Dim sql As String
                sql = "SELECT COUNT(*) FROM [" & tablaRef & "] WHERE [" & campoRef & "] = " & EscapeValue(valorFK)

                Dim rs As Object
                Set rs = m_dbConn.Connection.Execute(sql)

                If rs(0) = 0 Then
                    errors.Add campo, "FK inválida: " & campo & "=" & valorFK & " no existe en " & tablaRef
                End If
                rs.Close
            End If
        End If
    Next i

    Set ValidateForeignKeys = errors
    Exit Function

ErrorHandler:
    Debug.Print "[ERR] ValidateForeignKeys: " & Err.Description
    Set ValidateForeignKeys = errors
End Function

' ══════════════════════════════════════════════════════════
' DESCUBRIMIENTO DINÁMICO DE ESTRUCTURA DE BD
' ══════════════════════════════════════════════════════════

Public Function GetAllTables() As Variant
    ' Retorna array con nombres de todas las tablas de la BD
    On Error GoTo ErrorHandler

    Dim rs As Object
    Set rs = m_dbConn.Connection.OpenSchema(20) ' adSchemaTables = 20

    Dim tables As Object
    Set tables = CreateObject("System.Collections.ArrayList")

    Do Until rs.EOF
        Dim tableType As String
        tableType = NzVariant(rs("TABLE_TYPE"))

        ' Solo tablas (no vistas, ni tablas del sistema)
        If tableType = "TABLE" Then
            Dim tableName As String
            tableName = rs("TABLE_NAME")

            ' Excluir tablas del sistema (MSys*)
            If Left$(tableName, 4) <> "MSys" And Left$(tableName, 1) <> "~" Then
                tables.Add tableName
            End If
        End If
        rs.MoveNext
    Loop

    rs.Close

    ' Convertir ArrayList a Array
    Dim result() As String
    If tables.Count > 0 Then
        ReDim result(tables.Count - 1)
        Dim i As Long
        For i = 0 To tables.Count - 1
            result(i) = tables(i)
        Next i
        GetAllTables = result
    Else
        GetAllTables = Array()
    End If
    Exit Function

ErrorHandler:
    Debug.Print "[ERR] GetAllTables: " & Err.Description
    GetAllTables = Array()
End Function

Public Function GetChildTablesForFK(ByVal fkFieldName As String) As Variant
    ' Descubre dinámicamente todas las tablas que tienen un campo FK específico
    ' y ese campo NO es autonumérico (son tablas hijas)
    ' Retorna Array de nombres de tablas hijas
    On Error GoTo ErrorHandler

    Dim allTables As Variant
    allTables = GetAllTables()

    Dim childTables As Object
    Set childTables = CreateObject("System.Collections.ArrayList")

    Dim i As Long
    For i = LBound(allTables) To UBound(allTables)
        Dim tableName As String
        tableName = allTables(i)

        ' Verificar si la tabla tiene el campo FK
        If TableHasField(tableName, fkFieldName) Then
            ' Verificar que el campo NO sea autonumérico (si es autonumérico, es tabla maestra)
            If Not IsFieldAutoNumeric(tableName, fkFieldName) Then
                childTables.Add tableName
            End If
        End If
    Next i

    ' Convertir ArrayList a Array
    Dim result() As String
    If childTables.Count > 0 Then
        ReDim result(childTables.Count - 1)
        For i = 0 To childTables.Count - 1
            result(i) = childTables(i)
        Next i
        GetChildTablesForFK = result
    Else
        GetChildTablesForFK = Array()
    End If
    Exit Function

ErrorHandler:
    Debug.Print "[ERR] GetChildTablesForFK: " & Err.Description
    GetChildTablesForFK = Array()
End Function

Public Function GetTableFields(ByVal tableName As String) As Variant
    ' Retorna Array de nombres de campos de una tabla
    On Error GoTo ErrorHandler

    Dim rs As Object
    Set rs = m_dbConn.Connection.OpenSchema(4, Array(Empty, Empty, tableName)) ' adSchemaColumns = 4

    Dim fields As Object
    Set fields = CreateObject("System.Collections.ArrayList")

    Do Until rs.EOF
        fields.Add rs("COLUMN_NAME")
        rs.MoveNext
    Loop

    rs.Close

    ' Convertir ArrayList a Array
    Dim result() As String
    If fields.Count > 0 Then
        ReDim result(fields.Count - 1)
        Dim i As Long
        For i = 0 To fields.Count - 1
            result(i) = fields(i)
        Next i
        GetTableFields = result
    Else
        GetTableFields = Array()
    End If
    Exit Function

ErrorHandler:
    Debug.Print "[ERR] GetTableFields: " & Err.Description
    GetTableFields = Array()
End Function

' ══════════════════════════════════════════════════════════
' FUNCIONES DE INTROSPECCIÓN (movidas desde modSchema)
' ══════════════════════════════════════════════════════════

Private Function GetFieldsExcept(ByVal tableName As String, ByVal fieldsToExclude As Variant) As String
    ' Retorna lista de campos de tabla EXCEPTO los especificados
    ' fieldsToExclude: String o Array de Strings
    On Error GoTo ErrHandler

    Dim rs As Object
    Set rs = m_dbConn.Connection.OpenSchema(4, Array(Empty, Empty, tableName)) ' adSchemaColumns = 4

    Dim campo As String, lista As String
    Dim i As Long, excluir As Boolean

    Do Until rs.EOF
        campo = rs("COLUMN_NAME")
        excluir = False

        If IsArray(fieldsToExclude) Then
            For i = LBound(fieldsToExclude) To UBound(fieldsToExclude)
                If StrComp(campo, fieldsToExclude(i), vbTextCompare) = 0 Then
                    excluir = True
                    Exit For
                End If
            Next i
        Else
            If Not IsEmpty(fieldsToExclude) Then
                excluir = (StrComp(campo, fieldsToExclude, vbTextCompare) = 0)
            End If
        End If

        If Not excluir Then
            lista = lista & "[" & campo & "],"
        End If

        rs.MoveNext
    Loop

    rs.Close
    If Len(lista) > 0 Then lista = Left$(lista, Len(lista) - 1)
    GetFieldsExcept = lista
    Exit Function

ErrHandler:
    Debug.Print "[ERR] GetFieldsExcept: " & Err.Description
    GetFieldsExcept = ""
End Function

Private Function GetAutoNumericFields(ByVal tableName As String) As Variant
    ' Detecta campos autonuméricos (GUID con GenGUID() o Integer con flags específicos)
    On Error GoTo ErrHandler

    Dim rs As Object
    Set rs = m_dbConn.Connection.OpenSchema(4, Array(Empty, Empty, tableName))

    Dim campos() As String
    Dim contador As Long
    Const adInteger = 3
    Const adGUID = 72

    Do Until rs.EOF
        Dim tipo As Long: tipo = CLng(rs("DATA_TYPE"))
        Dim flags As Long: flags = NzVariant(rs("COLUMN_FLAGS"))
        If VarType(flags) <> vbLong Then flags = 0&
        Dim defaultVal As String: defaultVal = NzVariant(rs("COLUMN_DEFAULT"))
        If VarType(defaultVal) <> vbString Then defaultVal = ""
        Dim nombre As String: nombre = rs("COLUMN_NAME")

        Dim esAutonum As Boolean
        esAutonum = ( _
            (tipo = adGUID And defaultVal = "GenGUID()") Or _
            (tipo = adInteger And (flags = 17 Or flags = 90)) _
        )

        If esAutonum Then
            ReDim Preserve campos(contador)
            campos(contador) = nombre
            contador = contador + 1
        End If

        rs.MoveNext
    Loop

    rs.Close
    If contador = 0 Then
        GetAutoNumericFields = Array()
    Else
        GetAutoNumericFields = campos
    End If
    Exit Function

ErrHandler:
    Debug.Print "[ERR] GetAutoNumericFields: " & Err.Description
    GetAutoNumericFields = Array()
End Function

Private Function ReplaceFieldInSQL(ByVal fieldList As String, ByVal fieldName As String, ByVal newValue As String) As String
    ' Reemplaza un campo en lista SQL con nuevo valor
    Dim partes() As String
    partes = Split(fieldList, ",")
    Dim i As Long
    For i = 0 To UBound(partes)
        If Trim$(partes(i)) = "[" & fieldName & "]" Then
            partes(i) = newValue
        End If
    Next i
    ReplaceFieldInSQL = Join(partes, ",")
End Function

Private Function CombineArrays(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    ' Combina dos arrays en uno solo
    If Not IsArray(arr1) Then arr1 = Array()
    If Not IsArray(arr2) Then arr2 = Array()

    Dim size1 As Long, size2 As Long
    If UBound(arr1) >= LBound(arr1) Then
        size1 = UBound(arr1) - LBound(arr1) + 1
    Else
        size1 = 0
    End If

    If UBound(arr2) >= LBound(arr2) Then
        size2 = UBound(arr2) - LBound(arr2) + 1
    Else
        size2 = 0
    End If

    If size1 = 0 And size2 = 0 Then
        CombineArrays = Array()
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(size1 + size2 - 1)

    Dim i As Long, idx As Long
    idx = 0
    For i = LBound(arr1) To UBound(arr1)
        result(idx) = arr1(i)
        idx = idx + 1
    Next i

    For i = LBound(arr2) To UBound(arr2)
        result(idx) = arr2(i)
        idx = idx + 1
    Next i

    CombineArrays = result
End Function

Private Function TableHasField(ByVal tableName As String, ByVal fieldName As String) As Boolean
    ' Verifica si una tabla tiene un campo específico
    On Error GoTo ErrorHandler

    Dim rs As Object
    Set rs = m_dbConn.Connection.OpenSchema(4, Array(Empty, Empty, tableName, fieldName)) ' adSchemaColumns = 4

    TableHasField = Not rs.EOF
    rs.Close
    Exit Function

ErrorHandler:
    TableHasField = False
End Function

Private Function IsFieldAutoNumeric(ByVal tableName As String, ByVal fieldName As String) As Boolean
    ' Verifica si un campo es autonumérico (GUID o Integer con flags específicos)
    On Error GoTo ErrorHandler

    Dim rs As Object
    Set rs = m_dbConn.Connection.OpenSchema(4, Array(Empty, Empty, tableName, fieldName))

    If Not rs.EOF Then
        Dim tipo As Long
        tipo = CLng(rs("DATA_TYPE"))
        Dim flags As Long
        flags = NzVariant(rs("COLUMN_FLAGS"))
        If VarType(flags) <> vbLong Then flags = 0&
        Dim defaultVal As String
        defaultVal = NzVariant(rs("COLUMN_DEFAULT"))
        If VarType(defaultVal) <> vbString Then defaultVal = ""

        Const adInteger = 3
        Const adGUID = 72

        IsFieldAutoNumeric = ( _
            (tipo = adGUID And defaultVal = "GenGUID()") Or _
            (tipo = adInteger And (flags = 17 Or flags = 90)) _
        )
    Else
        IsFieldAutoNumeric = False
    End If

    rs.Close
    Exit Function

ErrorHandler:
    IsFieldAutoNumeric = False
End Function
