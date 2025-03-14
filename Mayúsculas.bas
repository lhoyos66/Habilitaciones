Attribute VB_Name = "Mayúsculas"
Sub ConvertirMayusculasOptimizado(sheetName As String)
    Dim ws As Worksheet
    Dim rng As Range
    Dim datos As Variant
    Dim i As Long, j As Long

    ' Establecer la hoja utilizando el nombre proporcionado
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' Establecer el rango desde la fila 1 a la 5394 y desde la columna A a la columna L
    Set rng = ws.Range("A1:L5394")
    
    ' Copiar los datos del rango a una matriz
    datos = rng.Value
    
    ' Recorrer la matriz y convertir cada valor a mayúsculas si es texto
    For i = 1 To UBound(datos, 1)
        For j = 1 To UBound(datos, 2)
            If VarType(datos(i, j)) = vbString Then
                datos(i, j) = UCase(datos(i, j))
            End If
        Next j
    Next i
    
    ' Copiar los datos de vuelta al rango
    rng.Value = datos
    
    MsgBox "La conversión a mayúsculas ha sido completada."
End Sub

Sub EjecutarConvertirMayusculas()
    ' Llamar al procedimiento con el nombre de la hoja "PERSONAL HISTORICO"
    Call ConvertirMayusculasOptimizado("BD VEHICULOS")
    Call ConvertirMayusculasOptimizado("BD PERSONAL")
End Sub

