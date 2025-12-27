VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDBContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================
'@Description: Contexto de conexión a base de datos Access
'@Purpose: Encapsular conexión ADODB y operaciones básicas
'=========================================================
Option Explicit
'TODO: migrar **La parte de lectura de registros** de este sistema de conexiones, si es posible, al sistema que expone VBA y Excel: Workbook.Connections; Workbook.Quieries
' (Para escribir registros hay que seguir usando el método CreateObject("ADODB.Connection"); Haciendo la lectura con este método se puede visualizar la relación de tablas
' Desde la interfaz de Excel y ello permite Acceder a funcionalidades avanzadas de tablas dinámicas etc)
Private m_Cnn As Object ' ADODB.Connection

'@Description: Conecta a la base de datos Access
'@Param RutaBD: Ruta completa al archivo .mdb o .accdb
Public Sub Conectar(ByVal RutaBD As String)
    On Error GoTo ErrHandler
    
    ' Validaciones
    If Len(RutaBD) = 0 Then
        Err.Raise 5, "clsDBContext.Conectar", "Ruta de BD vacía"
    End If
    
    If Dir(RutaBD) = "" Then
        Err.Raise 5, "clsDBContext.Conectar", _
            "Archivo de BD no encontrado: " & vbCrLf & RutaBD
    End If
    
    ' Crear conexión
    Set m_Cnn = CreateObject("ADODB.Connection")
    
    ' Abrir conexión (intentar varios providers)
    On Error Resume Next
    m_Cnn.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & RutaBD
    If Err.Number <> 0 Then
        Err.Clear
        m_Cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & RutaBD
    End If
    If Err.Number <> 0 Then
        Err.Clear
        m_Cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & RutaBD
    End If
    On Error GoTo ErrHandler
    
    If Not EstaConectado Then
        Err.Raise 5, "clsDBContext.Conectar", _
            "No se pudo establecer conexión con la BD"
    End If
    
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, "clsDBContext.Conectar", _
        "Error al conectar a BD: " & Err.Description
End Sub

'@Description: Verifica si hay conexión activa
'@Returns: True si conectado, False si no
Public Function EstaConectado() As Boolean
    EstaConectado = Not m_Cnn Is Nothing And m_Cnn.State = 1
End Function

'@Description: Cierra la conexión a la BD
Public Sub Desconectar()
    On Error Resume Next
    If Not m_Cnn Is Nothing Then
        If m_Cnn.State = 1 Then m_Cnn.Close
        Set m_Cnn = Nothing
    End If
    On Error GoTo 0
End Sub

'@Description: Abre un recordset con la consulta SQL proporcionada
'@Param Sql: Consulta SQL a ejecutar
'@Returns: ADODB.Recordset (como Object)
Public Function AbrirRecordset(ByVal sql As String) As Object
    On Error GoTo ErrHandler
    
    If sql = "" Then
        Err.Raise 5, "clsDBContext.AbrirRecordset", "SQL vacío"
    End If
    
    If Not EstaConectado Then
        Err.Raise 5, "clsDBContext.AbrirRecordset", "No hay conexión activa"
    End If
    
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, m_Cnn, 3, 1 ' adOpenStatic, adLockReadOnly
    
    Set AbrirRecordset = rs
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "clsDBContext.AbrirRecordset", _
        "Error al abrir recordset: " & Err.Description & vbCrLf & _
        "SQL: " & sql
End Function

'@Description: Ejecuta una consulta de acción (INSERT, UPDATE, DELETE)
'@Param Sql: Consulta SQL a ejecutar
'@Returns: Número de registros afectados
Public Function EjecutarComando(ByVal sql As String) As Long
    On Error GoTo ErrHandler
    
    If sql = "" Then
        Err.Raise 5, "clsDBContext.EjecutarComando", "SQL vacío"
    End If
    
    If Not EstaConectado Then
        Err.Raise 5, "clsDBContext.EjecutarComando", "No hay conexión activa"
    End If
    
    Dim recordsAffected As Long
    m_Cnn.Execute sql, recordsAffected
    EjecutarComando = recordsAffected
    Exit Function
    
ErrHandler:
    Err.Raise Err.Number, "clsDBContext.EjecutarComando", _
        "Error al ejecutar comando: " & Err.Description & vbCrLf & _
        "SQL: " & sql
End Function

'@Description: Obtiene la conexión subyacente (para operaciones avanzadas)
'@Returns: ADODB.Connection (como Object)
Public Function ObtenerConexion() As Object
    Set ObtenerConexion = m_Cnn
End Function

Private Sub Class_Terminate()
    Desconectar
End Sub
