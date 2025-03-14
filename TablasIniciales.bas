Attribute VB_Name = "TablasIniciales"
Sub CrearHojaResumen()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim nextRow As Long
    Dim dict As Object
    Dim key As String
    Dim puesto As String
    Dim documento As String

    ' Establecer la hoja origen
    Set wsOrigen = ThisWorkbook.Sheets("Personal")
    
    ' Crear la hoja destino
    On Error Resume Next
    Set wsDestino = ThisWorkbook.Sheets("Resumen")
    If wsDestino Is Nothing Then
        Set wsDestino = ThisWorkbook.Sheets.Add(After:=wsOrigen)
        wsDestino.Name = "Resumen"
    End If
    On Error GoTo 0
    
    ' Limpiar la hoja destino
    wsDestino.Cells.Clear
    
    ' Añadir encabezados a la hoja destino
    wsDestino.Range("A1:D1").Value = Array("Código de Proyecto", "Empresa", "Documento", "Servicio o Puesto")
    
    ' Inicializar el diccionario
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Obtener la última fila con datos en la hoja origen
    lastRow = wsOrigen.Cells(wsOrigen.Rows.Count, "A").End(xlUp).Row
    
    ' Recorrer las filas de la hoja origen y agregar los datos al diccionario
    For i = 2 To lastRow
        puesto = UCase(wsOrigen.Cells(i, 3).Value)
        documento = UCase(wsOrigen.Cells(i, 6).Value)
        key = wsOrigen.Cells(i, 1).Value & "|" & wsOrigen.Cells(i, 2).Value & "|" & documento & "|" & puesto
        If Not dict.Exists(key) Then
            dict.Add key, Array(wsOrigen.Cells(i, 1).Value, wsOrigen.Cells(i, 2).Value, documento, puesto)
        End If
    Next i
    
    ' Inicializar la siguiente fila en la hoja destino
    nextRow = 2
    
    ' Escribir los datos del diccionario en la hoja destino
    For Each key1 In dict.keys
        wsDestino.Cells(nextRow, 1).Value = dict(key1)(0) ' Código de Proyecto
        wsDestino.Cells(nextRow, 2).Value = dict(key1)(1) ' Empresa
        wsDestino.Cells(nextRow, 3).Value = dict(key1)(2) ' Documento
        wsDestino.Cells(nextRow, 4).Value = dict(key1)(3) ' Servicio o Puesto
        nextRow = nextRow + 1
    Next key1
    
    ' Formato de la hoja destino
    wsDestino.Columns("A:D").AutoFit
    
    ' Mensaje de finalización
    MsgBox "Hoja de resumen creada correctamente.", vbInformation
End Sub



Sub CrearHojaResumenVehiculos()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim nextRow As Long
    Dim dict As Object
    Dim key As String

    ' Establecer la hoja origen
    Set wsOrigen = ThisWorkbook.Sheets("Vehículos")
    
    ' Crear la hoja destino
    On Error Resume Next
    Set wsDestino = ThisWorkbook.Sheets("Resumen Vehículos")
    If wsDestino Is Nothing Then
        Set wsDestino = ThisWorkbook.Sheets.Add(After:=wsOrigen)
        wsDestino.Name = "Resumen Vehículos"
    End If
    On Error GoTo 0
    
    ' Limpiar la hoja destino
    wsDestino.Cells.Clear
    
    ' Añadir encabezados a la hoja destino
    wsDestino.Range("A1:E1").Value = Array("Código de Proyecto", "Empresa", "Servicio", "Tipo de Vehículo", "Requisito")
    
    ' Inicializar el diccionario
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Obtener la última fila con datos en la hoja origen
    lastRow = wsOrigen.Cells(wsOrigen.Rows.Count, "A").End(xlUp).Row
    
    ' Recorrer las filas de la hoja origen y agregar los datos al diccionario
    For i = 2 To lastRow
        key = wsOrigen.Cells(i, 1).Value & "|" & wsOrigen.Cells(i, 2).Value & "|" & UCase(wsOrigen.Cells(i, 3).Value) & "|" & UCase(wsOrigen.Cells(i, 5).Value) & "|" & UCase(wsOrigen.Cells(i, 6).Value)
        If Not dict.Exists(key) Then
            dict.Add key, Array(wsOrigen.Cells(i, 1).Value, wsOrigen.Cells(i, 2).Value, UCase(wsOrigen.Cells(i, 3).Value), UCase(wsOrigen.Cells(i, 5).Value), UCase(wsOrigen.Cells(i, 6).Value))
        End If
    Next i
    
    ' Inicializar la siguiente fila en la hoja destino
    nextRow = 2
    
    ' Escribir los datos del diccionario en la hoja destino
    For Each key1 In dict.keys
        wsDestino.Cells(nextRow, 1).Value = dict(key1)(0) ' Código de Proyecto
        wsDestino.Cells(nextRow, 2).Value = dict(key1)(1) ' Empresa
        wsDestino.Cells(nextRow, 3).Value = dict(key1)(2) ' Servicio
        wsDestino.Cells(nextRow, 4).Value = dict(key1)(3) ' Tipo de Vehículo
        wsDestino.Cells(nextRow, 5).Value = dict(key1)(4) ' Requisito
        nextRow = nextRow + 1
    Next key1
    
    ' Formato de la hoja destino
    wsDestino.Columns("A:E").AutoFit
    
    ' Mensaje de finalización
    MsgBox "Hoja de resumen de vehículos creada correctamente.", vbInformation
End Sub


Sub AgruparDocumentosPorServicio()
    Dim wsOriginal As Worksheet
    Dim wsResumen As Worksheet
    Dim lastRow As Long
    Dim summaryRow As Long
    Dim i As Long
    Dim dict As Object
    Dim key As Variant
    Dim uniqueDict As Object

    ' Establecer las hojas de trabajo
    On Error Resume Next
    Set wsOriginal = ThisWorkbook.Sheets("Resumen") ' Cambia "HojaOriginal" al nombre de tu hoja original
    On Error GoTo 0
    If wsOriginal Is Nothing Then
        MsgBox "La hoja 'Resumen' no existe.", vbExclamation
        Exit Sub
    End If
    
    Set wsResumen = ThisWorkbook.Sheets.Add
    wsResumen.Name = "ResumenAgrupado"
    
    ' Crear un diccionario para almacenar los datos agrupados
    Set dict = CreateObject("Scripting.Dictionary")
    Set uniqueDict = CreateObject("Scripting.Dictionary")

    ' Obtener la última fila con datos en la hoja original
    lastRow = wsOriginal.Cells(wsOriginal.Rows.Count, "A").End(xlUp).Row

    ' Recorrer las filas de la hoja original
    For i = 2 To lastRow
        Dim Servicio As String
        Dim documento As String
        Dim combinedKey As String
        
        Servicio = wsOriginal.Cells(i, 4).Value ' Asumiendo que "Servicio o Puesto" está en la columna D
        documento = wsOriginal.Cells(i, 3).Value ' Asumiendo que "Documento" está en la columna C
        combinedKey = Servicio & "|" & documento ' Usamos un delimitador para crear una clave única

        If Not uniqueDict.Exists(combinedKey) Then
            uniqueDict.Add combinedKey, Nothing
            If Not dict.Exists(Servicio) Then
                Set dict(Servicio) = New Collection
            End If
            dict(Servicio).Add documento
        End If
    Next i

    ' Escribir los datos agrupados en la hoja de resumen
    summaryRow = 1
    wsResumen.Cells(summaryRow, 1).Value = "Servicio o Puesto"
    wsResumen.Cells(summaryRow, 2).Value = "Documento"

    For Each key In dict.keys
        For i = 1 To dict(key).Count
            summaryRow = summaryRow + 1
            wsResumen.Cells(summaryRow, 1).Value = key
            wsResumen.Cells(summaryRow, 2).Value = dict(key)(i)
        Next i
    Next key

    ' Ajustar el ancho de las columnas
    wsResumen.Columns("A:B").AutoFit

    ' Limpiar los diccionarios
    Set dict = Nothing
    Set uniqueDict = Nothing

    MsgBox "Documentos agrupados por Servicio o Puesto creados con éxito en la hoja 'ResumenAgrupado'", vbInformation
End Sub

Sub AgruparRequisitosPorServicioYTipoDeVehiculo()
    Dim wsOriginal As Worksheet
    Dim wsResumen As Worksheet
    Dim lastRow As Long
    Dim summaryRow As Long
    Dim i As Long
    Dim dict As Object
    Dim uniqueDict As Object
    Dim key As Variant

    ' Establecer las hojas de trabajo
    On Error Resume Next
    Set wsOriginal = ThisWorkbook.Sheets("Resumen Vehículos") ' Cambia "Resumen Vehículos" al nombre de tu hoja original
    On Error GoTo 0
    If wsOriginal Is Nothing Then
        MsgBox "La hoja 'Resumen Vehículos' no existe.", vbExclamation
        Exit Sub
    End If
    
    Set wsResumen = ThisWorkbook.Sheets.Add
    wsResumen.Name = "ResumenAgrupadoVehículos"

    ' Crear diccionarios para almacenar los datos agrupados y las combinaciones únicas
    Set dict = CreateObject("Scripting.Dictionary")
    Set uniqueDict = CreateObject("Scripting.Dictionary")

    ' Obtener la última fila con datos en la hoja original
    lastRow = wsOriginal.Cells(wsOriginal.Rows.Count, "A").End(xlUp).Row

    ' Recorrer las filas de la hoja original
    For i = 2 To lastRow
        Dim Servicio As String
        Dim tipoVehiculo As String
        Dim requisito As String
        Dim combinedKey As String
        Dim uniqueKey As String
        
        Servicio = wsOriginal.Cells(i, 3).Value ' Asumiendo que "Servicio" está en la columna C
        tipoVehiculo = wsOriginal.Cells(i, 4).Value ' Asumiendo que "Tipo de Vehículo" está en la columna D
        requisito = wsOriginal.Cells(i, 5).Value ' Asumiendo que "Requisito" está en la columna E
        combinedKey = Servicio & " - " & tipoVehiculo
        uniqueKey = combinedKey & " - " & requisito ' Crear una clave única para cada combinación

        If Not uniqueDict.Exists(uniqueKey) Then
            uniqueDict.Add uniqueKey, Nothing
            If Not dict.Exists(combinedKey) Then
                Set dict(combinedKey) = New Collection
            End If
            dict(combinedKey).Add requisito
        End If
    Next i

    ' Escribir los datos agrupados en la hoja de resumen
    summaryRow = 1
    wsResumen.Cells(summaryRow, 1).Value = "Servicio"
    wsResumen.Cells(summaryRow, 2).Value = "Tipo de Vehículo"
    wsResumen.Cells(summaryRow, 3).Value = "Requisitos"

    For Each key In dict.keys
        For i = 1 To dict(key).Count
            summaryRow = summaryRow + 1
            Dim keys() As String
            keys = Split(key, " - ")
            wsResumen.Cells(summaryRow, 1).Value = keys(0)
            wsResumen.Cells(summaryRow, 2).Value = keys(1)
            wsResumen.Cells(summaryRow, 3).Value = dict(key)(i)
        Next i
    Next key

    ' Ajustar el ancho de las columnas
    wsResumen.Columns("A:C").AutoFit

    ' Limpiar los diccionarios
    Set dict = Nothing
    Set uniqueDict = Nothing

    MsgBox "Requisitos agrupados por Servicio y Tipo de Vehículo creados con éxito en la hoja 'ResumenAgrupadoVehículos'", vbInformation
End Sub


Sub CrearDocumentoPorPuesto()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimoRenglon As Long
    Dim dict As Object
    Dim dictUnico As Object
    Dim puesto As String
    Dim documento As String
    Dim i As Long
    Dim colDocumentos As Collection
    Dim key As Variant
    Dim item As Variant

    ' Establecer la hoja de origen
    Set wsOrigen = ThisWorkbook.Sheets("DASHBOARD PERSONAL")
    ultimoRenglon = wsOrigen.Cells(wsOrigen.Rows.Count, "C").End(xlUp).Row

    ' Crear un diccionario para almacenar los documentos por puesto
    Set dict = CreateObject("Scripting.Dictionary")
    ' Crear un diccionario para asegurar la unicidad
    Set dictUnico = CreateObject("Scripting.Dictionary")

    ' Recorrer la hoja de origen para llenar el diccionario
    For i = 11 To ultimoRenglon
        puesto = wsOrigen.Cells(i, 3).Value
        documento = wsOrigen.Cells(i, 6).Value
        If Not dict.Exists(puesto) Then
            Set colDocumentos = New Collection
            dict.Add puesto, colDocumentos
        End If
        ' Solo agregar si la combinación de puesto y documento es única
        If Not dictUnico.Exists(puesto & "|" & documento) Then
            dict(puesto).Add documento
            dictUnico.Add puesto & "|" & documento, Nothing
        End If
    Next i

    ' Crear la hoja de destino
    On Error Resume Next
    Set wsDestino = ThisWorkbook.Sheets("DOCUMENTOSXPUESTO")
    On Error GoTo 0
    If wsDestino Is Nothing Then
        Set wsDestino = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDestino.Name = "DOCUMENTOSXPUESTO"
    End If
    wsDestino.Cells.Clear

    ' Escribir los datos en la hoja de destino
    wsDestino.Cells(1, 1).Value = "PUESTO"
    wsDestino.Cells(1, 2).Value = "DOCUMENTO"

    i = 2
    For Each key In dict.keys
        For Each item In dict(key)
            wsDestino.Cells(i, 1).Value = key
            wsDestino.Cells(i, 2).Value = item
            i = i + 1
        Next item
    Next key

    MsgBox "La hoja 'DOCUMENTOSXPUESTO' ha sido creada exitosamente."
End Sub

