Attribute VB_Name = "modUtils"
' modUtils.bas
Option Explicit

Public Function NzVBA(ByVal expr As Variant, Optional ByVal valorPorDefecto As Variant = "") As Variant
    If IsNull(expr) Or IsEmpty(expr) Or VarType(expr) = vbError Or expr Is Nothing Then
        NzVBA = valorPorDefecto
    Else
        NzVBA = expr
    End If
End Function

' ============================================================================
' EXPORTAR/IMPORTAR DATOS
' ============================================================================

Public Function ExportToExcel(ByVal tableName As String, ByVal data As Object, ByVal filePath As String) As Boolean
    On Error GoTo ErrHandler

    If data Is Nothing Then
        Debug.Print "[ERR] ExportToExcel: data es Nothing"
        ExportToExcel = False
        Exit Function
    End If

    If data.Count = 0 Then
        Debug.Print "[WARN] ExportToExcel: No hay datos para exportar"
        ExportToExcel = True
        Exit Function
    End If

    ' Late binding a Excel
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False

    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    xlSheet.Name = Left$(tableName, 31) ' Max 31 caracteres para nombre de hoja

    ' Obtener campos del primer registro
    Dim firstRec As Object
    Set firstRec = data(0)

    Dim campos() As Variant
    ReDim campos(firstRec.Count - 1)
    Dim i As Long
    i = 0
    Dim key As Variant
    For Each key In firstRec.Keys
        campos(i) = key
        i = i + 1
    Next key

    ' Escribir encabezados
    Dim col As Long
    For col = 0 To UBound(campos)
        xlSheet.Cells(1, col + 1).Value = campos(col)
    Next col

    ' Escribir datos
    Dim row As Long
    Dim rec As Object
    For row = 0 To data.Count - 1
        Set rec = data(row)
        For col = 0 To UBound(campos)
            xlSheet.Cells(row + 2, col + 1).Value = NzVBA(rec(campos(col)), "")
        Next col
    Next row

    ' Autoajustar columnas
    xlSheet.Columns.AutoFit

    ' Guardar y cerrar
    xlBook.SaveAs filePath
    xlBook.Close False
    xlApp.Quit

    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing

    Debug.Print "[OK] Exportados " & data.Count & " registros a " & filePath
    ExportToExcel = True
    Exit Function

ErrHandler:
    Debug.Print "[ERR] ExportToExcel: " & Err.Description
    On Error Resume Next
    If Not xlBook Is Nothing Then xlBook.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    ExportToExcel = False
End Function

Public Function ExportToCSV(ByVal tableName As String, ByVal data As Object, ByVal filePath As String) As Boolean
    On Error GoTo ErrHandler

    If data Is Nothing Then
        Debug.Print "[ERR] ExportToCSV: data es Nothing"
        ExportToCSV = False
        Exit Function
    End If

    If data.Count = 0 Then
        Debug.Print "[WARN] ExportToCSV: No hay datos para exportar"
        ExportToCSV = True
        Exit Function
    End If

    Dim fso As Object
    Dim txtFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile(filePath, True)

    ' Obtener campos del primer registro
    Dim firstRec As Object
    Set firstRec = data(0)

    Dim campos() As Variant
    ReDim campos(firstRec.Count - 1)
    Dim i As Long
    i = 0
    Dim key As Variant
    For Each key In firstRec.Keys
        campos(i) = key
        i = i + 1
    Next key

    ' Escribir encabezados
    Dim linea As String
    linea = ""
    For i = 0 To UBound(campos)
        linea = linea & EscapeCSV(CStr(campos(i))) & ","
    Next i
    If Len(linea) > 0 Then linea = Left$(linea, Len(linea) - 1)
    txtFile.WriteLine linea

    ' Escribir datos
    Dim rec As Object
    Dim j As Long
    For j = 0 To data.Count - 1
        Set rec = data(j)
        linea = ""
        For i = 0 To UBound(campos)
            linea = linea & EscapeCSV(CStr(NzVBA(rec(campos(i)), ""))) & ","
        Next i
        If Len(linea) > 0 Then linea = Left$(linea, Len(linea) - 1)
        txtFile.WriteLine linea
    Next j

    txtFile.Close
    Set txtFile = Nothing
    Set fso = Nothing

    Debug.Print "[OK] Exportados " & data.Count & " registros a " & filePath
    ExportToCSV = True
    Exit Function

ErrHandler:
    Debug.Print "[ERR] ExportToCSV: " & Err.Description
    On Error Resume Next
    If Not txtFile Is Nothing Then txtFile.Close
    ExportToCSV = False
End Function

Public Function ImportFromExcel(ByVal filePath As String, Optional ByVal sheetName As String = "") As Object
    On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        Debug.Print "[ERR] ImportFromExcel: Archivo no existe: " & filePath
        Set ImportFromExcel = Nothing
        Exit Function
    End If

    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False

    Set xlBook = xlApp.Workbooks.Open(filePath)

    ' Usar primera hoja si no se especifica
    If sheetName = "" Then
        Set xlSheet = xlBook.Worksheets(1)
    Else
        Set xlSheet = xlBook.Worksheets(sheetName)
    End If

    ' Leer encabezados (primera fila)
    Dim lastCol As Long
    lastCol = xlSheet.Cells(1, xlSheet.Columns.Count).End(-4159).Column ' xlToLeft = -4159

    Dim campos() As String
    ReDim campos(lastCol - 1)
    Dim col As Long
    For col = 1 To lastCol
        campos(col - 1) = CStr(xlSheet.Cells(1, col).Value)
    Next col

    ' Leer datos
    Dim lastRow As Long
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, 1).End(-4162).row ' xlUp = -4162

    Dim result As Object
    Set result = CreateObject("System.Collections.ArrayList")

    Dim row As Long
    Dim rec As Object
    Dim valor As Variant

    For row = 2 To lastRow
        Set rec = CreateObject("Scripting.Dictionary")
        For col = 1 To lastCol
            valor = xlSheet.Cells(row, col).Value
            rec(campos(col - 1)) = NzVBA(valor, "")
        Next col
        result.Add rec
    Next row

    xlBook.Close False
    xlApp.Quit

    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing

    Debug.Print "[OK] Importados " & result.Count & " registros desde " & filePath
    Set ImportFromExcel = result
    Exit Function

ErrHandler:
    Debug.Print "[ERR] ImportFromExcel: " & Err.Description
    On Error Resume Next
    If Not xlBook Is Nothing Then xlBook.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set ImportFromExcel = Nothing
End Function

Public Function ImportFromCSV(ByVal filePath As String) As Object
    On Error GoTo ErrHandler

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(filePath) Then
        Debug.Print "[ERR] ImportFromCSV: Archivo no existe: " & filePath
        Set ImportFromCSV = Nothing
        Exit Function
    End If

    Dim txtFile As Object
    Set txtFile = fso.OpenTextFile(filePath, 1) ' ForReading = 1

    ' Leer encabezados
    If txtFile.AtEndOfStream Then
        Debug.Print "[WARN] ImportFromCSV: Archivo vacío"
        Set ImportFromCSV = CreateObject("System.Collections.ArrayList")
        txtFile.Close
        Exit Function
    End If

    Dim lineaEncabezados As String
    lineaEncabezados = txtFile.ReadLine

    Dim campos() As String
    campos = ParseCSVLine(lineaEncabezados)

    ' Leer datos
    Dim result As Object
    Set result = CreateObject("System.Collections.ArrayList")

    Dim linea As String
    Dim valores() As String
    Dim rec As Object
    Dim i As Long

    Do Until txtFile.AtEndOfStream
        linea = txtFile.ReadLine
        If Len(Trim$(linea)) > 0 Then
            valores = ParseCSVLine(linea)
            Set rec = CreateObject("Scripting.Dictionary")
            For i = 0 To UBound(campos)
                If i <= UBound(valores) Then
                    rec(campos(i)) = valores(i)
                Else
                    rec(campos(i)) = ""
                End If
            Next i
            result.Add rec
        End If
    Loop

    txtFile.Close
    Set txtFile = Nothing
    Set fso = Nothing

    Debug.Print "[OK] Importados " & result.Count & " registros desde " & filePath
    Set ImportFromCSV = result
    Exit Function

ErrHandler:
    Debug.Print "[ERR] ImportFromCSV: " & Err.Description
    On Error Resume Next
    If Not txtFile Is Nothing Then txtFile.Close
    Set ImportFromCSV = Nothing
End Function

' ============================================================================
' FUNCIONES AUXILIARES
' ============================================================================

Private Function EscapeCSV(ByVal valor As String) As String
    ' Escapar comillas y envolver en comillas si contiene coma, comilla o salto de línea
    If InStr(1, valor, ",") > 0 Or InStr(1, valor, """") > 0 Or InStr(1, valor, vbCrLf) > 0 Or InStr(1, valor, vbLf) > 0 Then
        EscapeCSV = """" & Replace(valor, """", """""") & """"
    Else
        EscapeCSV = valor
    End If
End Function

Private Function ParseCSVLine(ByVal linea As String) As String()
    Dim result() As String
    Dim tempArray As Object
    Set tempArray = CreateObject("System.Collections.ArrayList")

    Dim i As Long
    Dim campo As String
    Dim enComillas As Boolean
    Dim char As String

    i = 1
    campo = ""
    enComillas = False

    Do While i <= Len(linea)
        char = Mid$(linea, i, 1)

        If char = """" Then
            If enComillas Then
                ' Verificar si es comilla doble (escape)
                If i < Len(linea) And Mid$(linea, i + 1, 1) = """" Then
                    campo = campo & """"
                    i = i + 1
                Else
                    enComillas = False
                End If
            Else
                enComillas = True
            End If
        ElseIf char = "," And Not enComillas Then
            tempArray.Add campo
            campo = ""
        Else
            campo = campo & char
        End If

        i = i + 1
    Loop

    ' Añadir último campo
    tempArray.Add campo

    ' Convertir ArrayList a Array
    ReDim result(tempArray.Count - 1)
    For i = 0 To tempArray.Count - 1
        result(i) = CStr(tempArray(i))
    Next i

    ParseCSVLine = result
End Function
