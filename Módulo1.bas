Attribute VB_Name = "Módulo1"
Sub DepurarDuplicados()

    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim lastRow As Long
    Dim dict As Object
    Dim key As Variant
    Dim i As Long
    Dim rowCount As Long
    Dim item As Variant
    Dim emissionDate As Variant
    Dim validityDate As Variant
    
    ' Set the sheet "PERSONAL DOCUMENTOS"
    Set ws = ThisWorkbook.Sheets("PERSONAL DOCUMENTOS")
    
    ' Get the last row in the sheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Create a dictionary to store duplicates
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Create a new worksheet for the results
    Set newWs = ThisWorkbook.Sheets.Add
    newWs.Name = "Depurados Duplicados"
    
    ' Add headers to the new sheet
    newWs.Cells(1, 1).Value = "DOC IDENTIDAD"
    newWs.Cells(1, 2).Value = "DOCUMENTO"
    newWs.Cells(1, 3).Value = "EMISION"
    newWs.Cells(1, 4).Value = "VIGENCIA"
    
    rowCount = 2 ' Start row for results
    
    ' Loop through data from row 2 to the last row
    For i = 2 To lastRow
        ' Combine DOC IDENTIDAD and DOCUMENTO as key, treating DOC IDENTIDAD as text
        key = CStr(ws.Cells(i, 1).Value) & "|" & ws.Cells(i, 2).Value
        
        ' Check if DOC IDENTIDAD and DOCUMENTO are not empty
        If CStr(ws.Cells(i, 1).Value) <> "" And ws.Cells(i, 2).Value <> "" Then
            ' Get emission and validity dates
            emissionDate = ws.Cells(i, 3).Value
            validityDate = ws.Cells(i, 4).Value
            obsDocumento = ws.Cells(i, 5).Value
            
            ' Check if the key exists in the dictionary
            If dict.exists(key) Then
                ' Append the emission and validity dates by concatenating them
                dict(key) = dict(key) & "|" & emissionDate & "," & validityDate & "," & obsDocumento
            Else
                ' Create a new entry in the dictionary
                dict(key) = emissionDate & "," & validityDate & "," & obsDocumento
            End If
        End If
    Next i
    
    ' Write the results to the new sheet
    For Each key In dict.keys
        ' Split each key and its corresponding dates
        Dim parts As Variant
        Dim dates As Variant
        Dim j As Long
        
        parts = Split(dict(key), "|")
        
        ' If there are more than one date for the same DOC IDENTIDAD and DOCUMENTO, list them
        If UBound(parts) > 0 Then
            For j = 0 To UBound(parts)
                dates = Split(parts(j), ",")
                newWs.Cells(rowCount, 1).Value = "'" & Split(key, "|")(0) ' DOC IDENTIDAD
                newWs.Cells(rowCount, 2).Value = Split(key, "|")(1) ' DOCUMENTO
                newWs.Cells(rowCount, 3).Value = dates(0) ' EMISION
                newWs.Cells(rowCount, 4).Value = dates(1) ' VIGENCIA
                newWs.Cells(rowCount, 5).Value = dates(2) ' OBSERVACION
                rowCount = rowCount + 1
            Next j
        End If
    Next key
    
    ' Completion message
    MsgBox "Depuración completada. Se encontraron registros duplicados con todas las fechas de emisión y vigencia.", vbInformation

End Sub


Sub ProcesarDepuradosDuplicados()

    Dim wsDepurados As Worksheet
    Dim wsPersonal As Worksheet
    Dim lastRowDepurados As Long
    Dim lastRowPersonal As Long
    Dim i As Long
    Dim id As String
    Dim documento As String
    Dim j As Long
    Dim rngDelete As Range
    Dim cell As Range

    ' Establecer las hojas de trabajo
    Set wsDepurados = ThisWorkbook.Sheets("Depurados Duplicados")
    Set wsPersonal = ThisWorkbook.Sheets("PERSONAL DOCUMENTOS")
    
    ' Obtener la última fila de ambas hojas
    lastRowDepurados = wsDepurados.Cells(wsDepurados.Rows.Count, "A").End(xlUp).Row
    lastRowPersonal = wsPersonal.Cells(wsPersonal.Rows.Count, "A").End(xlUp).Row

    ' PRIMERA PARTE: Eliminar las filas que en columna F de Depurados Duplicados tengan un 1
    For i = 2 To lastRowDepurados
        If wsDepurados.Cells(i, "F").Value = 1 Then
            id = wsDepurados.Cells(i, "A").Value ' DOC IDENTIDAD como texto
            documento = wsDepurados.Cells(i, "B").Value ' DOCUMENTO
            
            ' Buscar todas las filas coincidentes en PERSONAL DOCUMENTOS y preparar la eliminación
            For j = lastRowPersonal To 2 Step -1
                If wsPersonal.Cells(j, "A").Value = id And wsPersonal.Cells(j, "B").Value = documento Then
                    If rngDelete Is Nothing Then
                        Set rngDelete = wsPersonal.Rows(j)
                    Else
                        Set rngDelete = Union(rngDelete, wsPersonal.Rows(j))
                    End If
                End If
            Next j
        End If
    Next i

    ' Eliminar las filas si hay filas seleccionadas
    If Not rngDelete Is Nothing Then
        rngDelete.Delete
    End If

    ' Actualizar la última fila de PERSONAL DOCUMENTOS después de eliminar filas
    lastRowPersonal = wsPersonal.Cells(wsPersonal.Rows.Count, "A").End(xlUp).Row

    ' SEGUNDA PARTE: Copiar las filas que en columna F de Depurados Duplicados tengan un 2
    For i = 2 To lastRowDepurados
        If wsDepurados.Cells(i, "F").Value = 2 Then
            ' Copiar columnas A, B, C, D y E de Depurados Duplicados a PERSONAL DOCUMENTOS
            lastRowPersonal = lastRowPersonal + 1
            wsPersonal.Cells(lastRowPersonal, "A").Value = "'" & wsDepurados.Cells(i, "A").Value ' DOC IDENTIDAD
            wsPersonal.Cells(lastRowPersonal, "B").Value = wsDepurados.Cells(i, "B").Value ' DOCUMENTO
            wsPersonal.Cells(lastRowPersonal, "C").Value = wsDepurados.Cells(i, "C").Value ' EMISION
            wsPersonal.Cells(lastRowPersonal, "D").Value = wsDepurados.Cells(i, "D").Value ' VIGENCIA
            wsPersonal.Cells(lastRowPersonal, "E").Value = wsDepurados.Cells(i, "E").Value ' OTRA INFORMACION (suponiendo que esta está en la columna E)
        End If
    Next i

    ' Mensaje de finalización
    MsgBox "Proceso completado. Filas eliminadas y copiadas según la columna F de la hoja Depurados Duplicados.", vbInformation

End Sub


