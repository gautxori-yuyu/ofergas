VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOferta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_dao As clsGenericDAO
Private m_errors As Object ' Scripting.Dictionary

Private Sub Class_Initialize()
    Set m_errors = CreateObject("Scripting.Dictionary")
End Sub

Public Property Set GenericDAO(ByVal dao As clsGenericDAO)
    Set m_dao = dao
End Property

' ——————————————————————————————————————————————————————
' CREA OFERTA COMPLETA (maestra + todas las tablas hijas)
' data: Dictionary con tabla => datos
'   - OfertasDatosGenerales: Dictionary (1 registro)
'   - OfertasManoObra: ArrayList de Dictionary (1:N)
' Devuelve OFER_ID nuevo si éxito, "" si fallo
' ——————————————————————————————————————————————————————
Public Function CrearOfertaCompleta(ByVal data As Object) As String
    On Error GoTo ErrorHandler
    
    ' Validar maestra
    If Not data.Exists("OfertasDatosGenerales") Then
        LogError "CrearOfertaCompleta", "Falta 'OfertasDatosGenerales' en data"
        Exit Function
    End If
    
    Dim db As Object
    Set db = m_dao.DBConnection.Connection
    
    ' === TRANSACCIÓN DAO (como tu Access) ===
    Dim ws As Object
    On Error Resume Next
    Set ws = m_dao.DBConnection.DAOWorkspace
    On Error GoTo ErrorHandler
    ws.BeginTrans
    
    Dim idOferta As String
    
    ' 1. Insertar maestra
    Dim maestra As Object: Set maestra = data("OfertasDatosGenerales")
    Dim resultMaestra As Object
    Set resultMaestra = m_dao.InsertRecord("OfertasDatosGenerales", maestra)
    If resultMaestra Is Nothing Then
        Err.Raise vbObjectError + 5001, "CrearOfertaCompleta", "Fallo en OfertasDatosGenerales"
    End If
    
    If resultMaestra.Exists("OFER_ID") Then
        idOferta = resultMaestra("OFER_ID")
    Else
        Err.Raise vbObjectError + 5002, "CrearOfertaCompleta", "No se recuperó OFER_ID"
    End If
    
    ' 2. Propagar OFER_ID a todas las hijas y insertar
    Dim tabla As Variant
    For Each tabla In data.Keys
        If tabla <> "OfertasDatosGenerales" Then
            Call PropagarOFER_ID(data(tabla), idOferta)
            Dim resultHija As Object
            Set resultHija = m_dao.InsertRecord(tabla, data(tabla))
            If resultHija Is Nothing Then
                Err.Raise vbObjectError + 5003, "CrearOfertaCompleta", "Fallo en " & tabla
            End If
        End If
    Next tabla
    
    ws.CommitTrans
    CrearOfertaCompleta = idOferta
    LogOK "CrearOfertaCompleta", "Oferta creada: " & idOferta
    Exit Function
    
ErrorHandler:
    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Rollback
        Set ws = Nothing
    End If
    
    Dim msg As String
    msg = "Error en CrearOfertaCompleta: " & Err.Description & " (Err=" & Err.Number & ")"
    LogError "CrearOfertaCompleta", msg
    
    ' Intentar limpiar maestra si se creó
    If idOferta <> "" Then
        On Error Resume Next
        m_dao.DeleteRecord "OfertasDatosGenerales", idOferta, "OFER_ID"
    End If
    
    Err.Raise Err.Number, "clsOferta.CrearOfertaCompleta", Err.Description
End Function

' ——————————————————————————————————————————————————————
' DUPLICA OFERTA (como tu DuplicarOferta)
' Devuelve nuevo OFER_ID
' ——————————————————————————————————————————————————————
Public Function DuplicarOferta(ByVal idOrigen As String, ByVal nuevoNumOferta As String) As String
    On Error GoTo ErrorHandler
    
    ' Validar unicidad (como tu Access)
    Dim sql As String
    sql = "SELECT COUNT(*) FROM OfertasDatosGenerales WHERE OFER_NUM_OFERTA = '" & Replace(nuevoNumOferta, "'", "''") & "'"
    Dim rs As Object
    Set rs = m_dao.DBConnection.Connection.Execute(sql)
    If CLng(rs(0)) > 0 Then
        rs.Close
        Err.Raise vbObjectError + 5010, "DuplicarOferta", "Ya existe oferta con ese número"
    End If
    rs.Close
    
    ' === TRANSACCIÓN ===
    Dim ws As Object
    Set ws = CreateObject("DAO.DBEngine.120").Workspaces(0)
    ws.BeginTrans
    
    ' 1. Deducir campos (simplificado, como tu DeducirCamposGrales)
    Dim dataMaestra As Object
    Set dataMaestra = CreateObject("Scripting.Dictionary")
    dataMaestra("OFER_NUM_OFERTA") = nuevoNumOferta
    dataMaestra("OFER_FECHA") = Date
    
    ' Copiar otros campos del origen
    sql = "SELECT OFER_CLIENTE, OFER_USUARIO_FINAL, PAIS_ID, GASE_ID " & _
          "FROM OfertasDatosGenerales WHERE OFER_ID = '" & idOrigen & "'"
    Set rs = m_dao.DBConnection.Connection.Execute(sql)
    If Not rs.EOF Then
        dataMaestra("OFER_CLIENTE") = NzVBA(rs("OFER_CLIENTE"), "")
        dataMaestra("OFER_USUARIO_FINAL") = NzVBA(rs("OFER_USUARIO_FINAL"), "")
        dataMaestra("PAIS_ID") = NzVBA(rs("PAIS_ID"), 0&)
        dataMaestra("GASE_ID") = NzVBA(rs("GASE_ID"), 0&)
    End If
    rs.Close
    
    ' 2. Insertar maestra
    Dim resultMaestra As Object
    Set resultMaestra = m_dao.InsertRecord("OfertasDatosGenerales", dataMaestra)
    If resultMaestra Is Nothing Then
        Err.Raise vbObjectError + 5011, "DuplicarOferta", "Fallo al insertar maestra"
    End If
    
    Dim idNuevo As String
    If resultMaestra.Exists("OFER_ID") Then
        idNuevo = resultMaestra("OFER_ID")
    Else
        Err.Raise vbObjectError + 5012, "DuplicarOferta", "No se recuperó OFER_ID"
    End If
    
    ' 3. Duplicar todas las tablas hijas
    Dim hijas As Variant
    hijas = Array( _
        "OfertasManoObra", "OfertasModelo", "OfertasOpciones", _
        "OfertasOtros", "OfertasRefrigeradores", "OfertasAccesorios", _
        "OfertasCabezal", "OfertasCalderines", "OfertasExtras", _
        "OfertasInstrumentacion" _
    )
    
    Dim t As Variant
    For Each t In hijas
        If Not m_dao.DuplicateRecord(CStr(t), idOrigen, idNuevo, "OFER_ID") Then
            Err.Raise vbObjectError + 5013, "DuplicarOferta", "Fallo al duplicar " & t
        End If
    Next t
    
    ws.CommitTrans
    DuplicarOferta = idNuevo
    LogOK "DuplicarOferta", "Oferta duplicada: " & idOrigen & " ? " & idNuevo
    Exit Function
    
ErrorHandler:
    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Rollback
        Set ws = Nothing
    End If
    
    ' Limpiar maestra si se creó
    If idNuevo <> "" Then
        On Error Resume Next
        m_dao.DeleteRecord "OfertasDatosGenerales", idNuevo, "OFER_ID"
    End If
    
    LogError "DuplicarOferta", Err.Description
    Err.Raise Err.Number, "clsOferta.DuplicarOferta", Err.Description
End Function

' ——————————————————————————————————————————————————————
' BORRAR OFERTA COMPLETA (como tu BorrarOferta)
' ——————————————————————————————————————————————————————
Public Sub BorrarOferta(ByVal idOferta As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Object
    Set ws = CreateObject("DAO.DBEngine.120").Workspaces(0)
    ws.BeginTrans
    
    Dim hijas As Variant
    hijas = Array( _
        "OfertasManoObra", "OfertasModelo", "OfertasOpciones", _
        "OfertasOtros", "OfertasRefrigeradores", "OfertasAccesorios", _
        "OfertasCabezal", "OfertasCalderines", "OfertasExtras", _
        "OfertasInstrumentacion" _
    )
    
    Dim t As Variant
    For Each t In hijas
        If Not m_dao.DeleteRecord(CStr(t), idOferta, "OFER_ID") Then
            Err.Raise vbObjectError + 5020, "BorrarOferta", "Fallo al borrar " & t
        End If
    Next t
    
    ' Borrar maestra
    If Not m_dao.DeleteRecord("OfertasDatosGenerales", idOferta, "OFER_ID") Then
        Err.Raise vbObjectError + 5021, "BorrarOferta", "Fallo al borrar maestra"
    End If
    
    ws.CommitTrans
    LogOK "BorrarOferta", "Oferta eliminada: " & idOferta
    Exit Sub
    
ErrorHandler:
    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Rollback
        Set ws = Nothing
    End If
    LogError "BorrarOferta", Err.Description
    Err.Raise Err.Number, "clsOferta.BorrarOferta", Err.Description
End Sub

' ——————————————————————————————————————————————————————
' AUXILIARES
' ——————————————————————————————————————————————————————
Private Sub PropagarOFER_ID(ByRef dataItem As Object, ByVal idOferta As String)
    Dim esLista As Boolean
    esLista = TypeName(dataItem) = "ArrayList"
    
    Dim lista As Object
    If esLista Then
        Set lista = dataItem
    Else
        Set lista = CreateObject("System.Collections.ArrayList")
        lista.Add dataItem
    End If
    
    Dim i As Long
    For i = 0 To lista.Count - 1
        Dim reg As Object: Set reg = lista(i)
        If Not reg.Exists("OFER_ID") Then
            reg.Add "OFER_ID", idOferta
        End If
    Next i
End Sub

Private Sub LogOK(ByVal method As String, ByVal msg As String)
    Debug.Print "[OK] clsOferta." & method & ": " & msg
End Sub

Private Sub LogError(ByVal method As String, ByVal msg As String)
    Debug.Print "[ERR] clsOferta." & method & ": " & msg
    ' Opcional: acumular en m_errors
    If Not m_errors.Exists(method) Then m_errors.Add method, Array()
    Dim arr() As Variant
    arr = m_errors(method)
    ReDim Preserve arr(0 To UBound(arr) + 1)
    arr(UBound(arr)) = msg
    m_errors(method) = arr
End Sub

' ——————————————————————————————————————————————————————
' ACCESO A ERRORES (para diagnóstico)
' ——————————————————————————————————————————————————————
Public Function GetErrors() As Object
    Set GetErrors = m_errors
End Function


