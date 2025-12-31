Attribute VB_Name = "modTestEjemplo"
' modTestEjemplo.bas — Ejemplo fiel a tu Access
Option Explicit
Sub Test_CrearOfertaCompleta()
    Dim conn As clsDBConnection: Set conn = New clsDBConnection
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
    Dim conn As clsDBConnection: Set conn = New clsDBConnection
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
