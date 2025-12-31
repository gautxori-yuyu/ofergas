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

Private m_dbConn As clsDBConnection  ' ? cambia de Object ? clsDBConnection

Public Property Set DBConnection(ByVal conn As clsDBConnection)
    Set m_dbConn = conn
End Property

Public Property Get DBConnection() As clsDBConnection
    Set DBConnection = m_dbConn
End Property
' ������������������������������������������������������
' INSERTA un registro (1:1) o m�ltiples (1:N)
' data:
'   - Dictionary ? 1 registro
'   - System.Collections.ArrayList de Dictionary ? N registros
' Devuelve un Dictionary con {campoAutonum => valorGenerado} del primer registro,
' o Nothing si fallo. Para m�ltiples registros, solo devuelve del primero (como Access).
' ������������������������������������������������������
Public Function InsertRecord(ByVal tableName As String, ByVal data As Object, _
                              Optional ByVal idFieldName As String = "OFER_ID") As Object
    On Error GoTo ErrorHandler
    
    Dim autoFields As Variant
    autoFields = CamposAutonumericos(m_dbConn.Connection, tableName)
    
    Dim camposDestino As String
    camposDestino = CamposSin(m_dbConn.Connection, tableName, autoFields)
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
        
        ' Solo recuperar autonum�ricos del PRIMER registro (como hace tu Access)
        If i = 0 Then
            Dim k As Long
            For k = LBound(autoFields) To UBound(autoFields)
                Dim campoAuto As String: campoAuto = autoFields(k)
                Dim valorGenerado As String
                valorGenerado = RecoveryIDForField(tableName, campoAuto, reg)
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

' ������������������������������������������������������
' DUPLICAR gen�rico (como tu DuplicarRegistroEnTablas)
' ������������������������������������������������������
Public Function DuplicateRecord(ByVal tableName As String, ByVal idOrigen As String, ByVal idNuevo As String, ByVal idFieldName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim autoFields As Variant
    autoFields = CamposAutonumericos(m_dbConn.Connection, tableName)
    
    Dim camposDestino As String
    camposDestino = CamposSin(m_dbConn.Connection, tableName, autoFields)
    If camposDestino = "" Then
        DuplicateRecord = True
        Exit Function
    End If
    
    Dim camposOrigen As String
    camposOrigen = camposDestino
    camposOrigen = ReemplazarCampoSQL(camposOrigen, idFieldName, "'" & idNuevo & "'")
    
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

' ������������������������������������������������������
' BORRAR gen�rico
' ������������������������������������������������������
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

' ������������������������������������������������������
' RECUPERAR valor de un campo autonum�rico espec�fico (no solo OFER_ID)
' ������������������������������������������������������
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

' ������������������������������������������������������
' FUNCIONES AUXILIARES
' ������������������������������������������������������

Private Function RecoveryIDForField(ByVal tableName As String, ByVal fieldName As String, ByVal data As Object) As String
    ' Estrategia: buscar por campos �nicos + orden descendente
    ' Para OFER_ID: OFER_NUM_OFERTA + OFER_FECHA
    If fieldName = "OFER_ID" And data.Exists("OFER_NUM_OFERTA") And data.Exists("OFER_FECHA") Then
        Dim num As String: num = EscapeString(data("OFER_NUM_OFERTA"))
        Dim fecha As String: fecha = EscapeDate(data("OFER_FECHA"))
        
        Dim sql As String
        sql = "SELECT [" & fieldName & "] FROM [" & tableName & "] " & _
              "WHERE [OFER_NUM_OFERTA] = " & num & " AND [OFER_FECHA] = " & fecha & " " & _
              "ORDER BY [OFER_FECHA] DESC"
        
        Dim rs As Object
        Set rs = m_dbConn.Connection.Execute(sql)
        If Not rs.EOF Then
            Dim idRaw As String: idRaw = NzVBA(rs(0), "")
            If InStr(idRaw, "{guid {") > 0 Then
                idRaw = Replace(Replace(idRaw, "{guid {", "{"), "}}", "}")
            End If
            RecoveryIDForField = idRaw
        End If
        rs.Close
    End If
    ' Para otros campos autonum�ricos (ej: OFMO_ID), implementar l�gica espec�fica si es necesaria
    ' Por ahora, solo OFER_ID tiene recuperaci�n definida
End Function
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
    autoFields = CamposAutonumericos(m_dbConn.Connection, tableName)

    Dim camposDestino As String
    camposDestino = CamposSin(m_dbConn.Connection, tableName, autoFields)

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
