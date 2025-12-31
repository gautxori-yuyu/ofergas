Attribute VB_Name = "modSchema"
' modSchema.bas — Introspección genérica (late binding ADO)
' Fiel a tu CamposSin, CamposAutonumericos, ReemplazarCampoSQL
Option Explicit

Public Function CamposSin(ByVal conn As Object, ByVal tbl As String, ByVal camposExcluir As Variant) As String
    On Error GoTo ErrHandler
    Dim rs As Object
    Set rs = conn.OpenSchema(4, Array(Empty, Empty, tbl)) ' adSchemaColumns
    
    Dim campo As String, lista As String
    Dim i As Long, excluir As Boolean
    
    Do Until rs.EOF
        campo = rs("COLUMN_NAME")
        excluir = False
        
        If IsArray(camposExcluir) Then
            For i = LBound(camposExcluir) To UBound(camposExcluir)
                If StrComp(campo, camposExcluir(i), vbTextCompare) = 0 Then
                    excluir = True
                    Exit For
                End If
            Next i
        Else
            excluir = (StrComp(campo, camposExcluir, vbTextCompare) = 0)
        End If
        
        If Not excluir Then
            lista = lista & "[" & campo & "],"
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
    If Len(lista) > 0 Then lista = Left$(lista, Len(lista) - 1)
    CamposSin = lista
    Exit Function
    
ErrHandler:
    Debug.Print "[ERR] CamposSin: " & Err.Description
    CamposSin = ""
End Function

Public Function CamposAutonumericos(ByVal conn As Object, ByVal nombreTabla As String) As Variant
    On Error GoTo ErrHandler
    Dim rs As Object
    Set rs = conn.OpenSchema(4, Array(Empty, Empty, nombreTabla))
    
    Dim campos() As String
    Dim contador As Long
    Const adInteger = 3
    Const adGUID = 72
    
    Do Until rs.EOF
        Dim tipo As Long: tipo = CLng(rs("DATA_TYPE"))
        Dim flags As Long: flags = NzVBA(rs("COLUMN_FLAGS"), 0&)
        Dim defaultVal As String: defaultVal = NzVBA(rs("COLUMN_DEFAULT"), "")
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
        CamposAutonumericos = Array()
    Else
        CamposAutonumericos = campos
    End If
    Exit Function
    
ErrHandler:
    Debug.Print "[ERR] CamposAutonumericos: " & Err.Description
    CamposAutonumericos = Array()
End Function

Public Function ReemplazarCampoSQL(ByVal lista As String, ByVal campo As String, ByVal nuevoValor As String) As String
    Dim partes() As String
    partes = Split(lista, ",")
    Dim i As Long
    For i = 0 To UBound(partes)
        If Trim$(partes(i)) = "[" & campo & "]" Then
            partes(i) = nuevoValor
        End If
    Next i
    ReemplazarCampoSQL = Join(partes, ",")
End Function

