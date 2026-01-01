Attribute VB_Name = "modDBCreator"
' modDBCreator.bas — Crea BD con tu estructura real
Option Explicit

Public Sub CrearBaseDeDatosEjemplo()
    Dim rutaBD As String
    rutaBD = ThisWorkbook.Path & "\Ofertas_Ejemplo.accdb"
    
    CrearBDVacia rutaBD
    
    Dim db As Object
    Set db = CreateObject("ADODB.Connection")
    db.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & rutaBD & ";"
    
    CrearTablaOfertasDatosGenerales db
    CrearTablaOfertasManoObra db
    ' ... crea el resto (abreviado)
    CrearTablaGases db
    CrearTablaPaises db
    
    InsertarDatosPrueba db
    
    db.Close
    MsgBox "[OK] BD creada: " & rutaBD
End Sub

Private Sub CrearBDVacia(ByVal fullPath As String)
    'se requiere:
    '“Microsoft Access Database Engine 2016 Redistributable (64 bits)”
    'Descargable desde: https://www.microsoft.com/en-us/download/details.aspx?id=54920
    On Error Resume Next: Kill fullPath: On Error GoTo 0
    Dim cat As Object: Set cat = CreateObject("ADOX.Catalog")
    cat.Create "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & fullPath & ";"
End Sub

Private Sub CrearTablaOfertasDatosGenerales(ByVal db As Object)
    db.Execute "CREATE TABLE OfertasDatosGenerales (" & _
        "OFER_ID GUID DEFAULT GenGUID() CONSTRAINT PK_OFER PRIMARY KEY, " & _
        "OFER_NUM_OFERTA TEXT(50), OFER_FECHA DATETIME, OFER_CLIENTE TEXT(100), " & _
        "OFER_USUARIO_FINAL TEXT(100), PAIS_ID LONG, OFER_NUM_CALCULO TEXT(50), " & _
        "OFER_ESTADO_OFERTA BYTE, OFER_OBSERVACIONES MEMO, GASE_ID LONG)"
End Sub

Private Sub CrearTablaOfertasManoObra(ByVal db As Object)
    db.Execute "CREATE TABLE OfertasManoObra (" & _
        "OFMO_ID GUID DEFAULT GenGUID() CONSTRAINT PK_OFMO PRIMARY KEY, OFER_ID GUID, OFMO_DESCRIPCION TEXT(100), OFMO_HORAS LONG)"
End Sub

Private Sub CrearTablaGases(ByVal db As Object)
    db.Execute "CREATE TABLE Gases (GASE_ID LONG CONSTRAINT PK_GASE PRIMARY KEY, GASE_DENOMINACION TEXT(50))"
    db.Execute "INSERT INTO Gases (GASE_ID, GASE_DENOMINACION) VALUES (1, 'Aire')"
    db.Execute "INSERT INTO Gases (GASE_ID, GASE_DENOMINACION) VALUES (2, 'Nitrógeno')"
End Sub

Private Sub CrearTablaPaises(ByVal db As Object)
    db.Execute "CREATE TABLE Paises (pais_id LONG CONSTRAINT PK_PAIS PRIMARY KEY, pais_nombre TEXT(50))"
    db.Execute "INSERT INTO Paises (pais_id, pais_nombre) VALUES (1, 'España')"
    db.Execute "INSERT INTO Paises (pais_id, pais_nombre) VALUES (2, 'Francia')"
End Sub

Private Sub InsertarDatosPrueba(ByVal db As Object)
    ' Insertar maestra (OFER_ID generado automáticamente)
    db.Execute "INSERT INTO OfertasDatosGenerales (OFER_NUM_OFERTA, OFER_FECHA, OFER_CLIENTE, GASE_ID, PAIS_ID) " & _
               "VALUES ('OF-2025-001', #2025/12/28#, 'Cliente A', 1, 1)"
    
    ' Recuperar GUID
    Dim rs As Object
    Set rs = db.Execute("SELECT OFER_ID FROM OfertasDatosGenerales WHERE OFER_NUM_OFERTA = 'OF-2025-001'")
    Dim id1 As String: id1 = NzVBA(rs(0), ""): rs.Close
    
    ' Insertar hija mínima
    'db.Execute "INSERT INTO OfertasManoObra (OFER_ID, OFMO_DESCRIPCION) VALUES (?, 'Instalación')", Array(id1)
    'La línea anterior no es válida es para DAO y aquí estamos usando ADO (con interpolacion segura)
'    db.Execute "INSERT INTO OfertasManoObra (OFER_ID, OFMO_DESCRIPCION, OFMO_HORAS) VALUES ('" & Replace(id1, "'", "''") & "', 'Instalación eléctrica', 8)"
    ' Opción B (con tu DAO):
    Dim dao As clsGenericDAO: Set dao = New clsGenericDAO
    Set dao.DBConnection.Connection = db
    Dim data As Object: Set data = CreateObject("Scripting.Dictionary")
    data("OFER_ID") = id1
    data("OFMO_DESCRIPCION") = "Instalación eléctrica"
    data("OFMO_HORAS") = 8
    dao.InsertRecord "OfertasManoObra", data
End Sub

