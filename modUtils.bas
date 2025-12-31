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
