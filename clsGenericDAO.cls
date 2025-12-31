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
' ——————————————————————————————————————————————————————
' INSERTA un registro (1:1) o múltiples (1:N)
' data:
'   - Dictionary ? 1 registro
'   - System.Collections.ArrayList de Dictionary ? N registros
' Devuelve un Dictionary con {campoAutonum => valorGenerado} del primer registro,
' o Nothing si fallo. Para múltiples registros, solo devuelve del primero (como Access).
' ——————————————————————————————————————————————————————
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
        
        ' Solo recuperar autonuméricos del PRIMER registro (como hace tu Access)
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

' ——————————————————————————————————————————————————————
' DUPLICAR genérico (como tu DuplicarRegistroEnTablas)
' ——————————————————————————————————————————————————————
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

' ——————————————————————————————————————————————————————
' BORRAR genérico
' ——————————————————————————————————————————————————————
Public Function DeleteRecord(ByVal tableName As String, ByVal idValue As String, ByVal idFieldName As String) As Boolean
    Dim sql As String
    sql = "DELETE FROM [" & tableName & "] WHERE [" & idFieldName & "] = '" & EscapeString(idValue) & "'"
    m_dbConn.Connection.Execute sql
    DeleteRecord = True
End Function

' ——————————————————————————————————————————————————————
' RECUPERAR valor de un campo autonumérico específico (no solo OFER_ID)
' ——————————————————————————————————————————————————————
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

' ——————————————————————————————————————————————————————
' FUNCIONES AUXILIARES
' ——————————————————————————————————————————————————————

Private Function RecoveryIDForField(ByVal tableName As String, ByVal fieldName As String, ByVal data As Object) As String
    ' Estrategia: buscar por campos únicos + orden descendente
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
    ' Para otros campos autonuméricos (ej: OFMO_ID), implementar lógica específica si es necesaria
    ' Por ahora, solo OFER_ID tiene recuperación definida
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
