Attribute VB_Name = "MantenimientoTablas"
Sub ActualizarTablaPorPuesto()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim tablaOrigen As ListObject
    Dim tablaDestino As ListObject
    Dim servicioPuesto As String
    Dim codigoProyecto As String
    Dim i As Long
    Dim j As Long

    ' Establecer las hojas de origen y destino
    Set wsOrigen = ThisWorkbook.Sheets("MASTER DOCUMENTOS PERSONAL")
    Set wsDestino = ThisWorkbook.Sheets("DOCUMENTO X PUESTO")
    
    ' Obtener el valor de la celda F3 y F2 en la hoja de destino
    servicioPuesto = wsDestino.Range("G3").Value
    codigoProyecto = wsDestino.Range("G2").Value

    ' Confirmar la ejecución de la macro
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("¿Está seguro que desea actualizar la tabla con el servicio o puesto de " & servicioPuesto & "?", vbYesNo + vbQuestion, "Confirmación")
    
    If respuesta = vbNo Then
        Exit Sub
    End If

    ' Establecer las tablas de origen y destino
    Set tablaOrigen = wsOrigen.ListObjects("DocumentosXPuesto")
    Set tablaDestino = wsDestino.ListObjects("DocumentoPorPuesto")

    ' Eliminar filas antiguas con el mismo SERVICIO O PUESTO en la tabla de destino
    For i = tablaDestino.ListRows.Count To 1 Step -1
        If tablaDestino.ListRows(i).Range(1, 1).Value = servicioPuesto And tablaDestino.ListRows(i).Range(1, 4).Value = codigoProyecto Then
            tablaDestino.ListRows(i).Delete
        End If
    Next i

    ' Recorrer la tabla de origen y copiar las filas que coincidan con el SERVICIO O PUESTO
    For i = 1 To tablaOrigen.ListRows.Count
        If tablaOrigen.DataBodyRange(i, 1).Value = servicioPuesto Then
            ' Agregar nueva fila a la tabla de destino
            With tablaDestino.ListRows.Add
                .Range(1, 1).Value = tablaOrigen.DataBodyRange(i, 1).Value ' PUESTO (SERVICIO O PUESTO)
                .Range(1, 2).Value = tablaOrigen.DataBodyRange(i, 2).Value ' DOCUMENTO
                .Range(1, 3).Value = "SI"
                .Range(1, 4).Value = codigoProyecto ' CÓDIGO PROYECTO
            End With
        End If
    Next i

    MsgBox "La tabla 'DocumentoPorPuesto' ha sido actualizada exitosamente.", vbInformation
End Sub


Sub ActualizarRequisitosPorVehiculo()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim tablaOrigen As ListObject
    Dim tablaDestino As ListObject
    Dim Servicio As String
    Dim tipo As String
    Dim codigoProyecto As String
    Dim i As Long
    Dim j As Long

    ' Establecer las hojas de origen y destino
    Set wsOrigen = ThisWorkbook.Sheets("MASTER REQUISITOS VEHICULO")
    Set wsDestino = ThisWorkbook.Sheets("REQUISITOS X VEHICULO")
    
    ' Obtener los valores de las celdas G3, G4 y G2 en la hoja de destino
    Servicio = wsDestino.Range("H3").Value
    tipo = wsDestino.Range("H4").Value
    codigoProyecto = wsDestino.Range("H2").Value

    ' Confirmar la ejecución de la macro
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("¿Está seguro que desea actualizar la tabla con el servicio " & Servicio & " y de tipo de vehículo " & tipo & "?", vbYesNo + vbQuestion, "Confirmación")
    
    If respuesta = vbNo Then
        Exit Sub
    End If

    ' Establecer las tablas de origen y destino
    Set tablaOrigen = wsOrigen.ListObjects("RequisitosPorVehículo")
    Set tablaDestino = wsDestino.ListObjects("RequisitosPorVehiculo")

    ' Eliminar filas antiguas con el mismo SERVICIO y TIPO en la tabla de destino
    For i = tablaDestino.ListRows.Count To 1 Step -1
        If tablaDestino.ListRows(i).Range(1, 1).Value = Servicio And _
           tablaDestino.ListRows(i).Range(1, 2).Value = tipo And _
           tablaDestino.ListRows(i).Range(1, 5).Value = codigoProyecto Then
            tablaDestino.ListRows(i).Delete
        End If
    Next i

    ' Recorrer la tabla de origen y copiar las filas que coincidan con el SERVICIO y TIPO
    For i = 1 To tablaOrigen.ListRows.Count
        If tablaOrigen.DataBodyRange(i, 1).Value = Servicio And _
           tablaOrigen.DataBodyRange(i, 2).Value = tipo Then
            ' Agregar nueva fila a la tabla de destino
            With tablaDestino.ListRows.Add
                .Range(1, 1).Value = tablaOrigen.DataBodyRange(i, 1).Value ' SERVICIO
                .Range(1, 2).Value = tablaOrigen.DataBodyRange(i, 2).Value ' TIPO
                .Range(1, 3).Value = tablaOrigen.DataBodyRange(i, 3).Value ' REQUISITO
                .Range(1, 4).Value = "SI"
                .Range(1, 5).Value = codigoProyecto ' CÓDIGO PROYECTO
            End With
        End If
    Next i

    MsgBox "La tabla 'RequisitosPorVehiculo' ha sido actualizada exitosamente.", vbInformation
End Sub

Sub InsertarPersonalProyecto()
    Dim wsPersonal As Worksheet
    Dim wsMasterPersonal As Worksheet
    Dim tblPersonalProyecto As ListObject
    Dim nuevaFila As ListRow
    Dim documento As String
    Dim ServicioOPuesto As String
    Dim Empresa As String
    Dim proyecto As String
    Dim nombreCompleto As String
    Dim rngDNI As Range
    Dim celdaDNI As Range
    Dim celdaExistente As Range
    Dim rngDocumentos As Range
    
    ' Definir las hojas
    Set wsPersonal = ThisWorkbook.Sheets("PERSONAL")
    Set wsMasterPersonal = ThisWorkbook.Sheets("MASTER PERSONAL")
    
    ' Obtener la tabla PersonalProyecto
    Set tblPersonalProyecto = wsPersonal.ListObjects("PersonalProyecto")
    
    ' Obtener los valores de las celdas
    documento = wsPersonal.Range("I6").Value
    ServicioOPuesto = wsPersonal.Range("I4").Value
    Empresa = wsPersonal.Range("I5").Value
    proyecto = wsPersonal.Range("I2").Value
    
    ' Buscar el nombre completo en la hoja MASTER PERSONAL
    Set rngDNI = wsMasterPersonal.ListObjects("PERSONAL").ListColumns("DNI").DataBodyRange
    Set celdaDNI = rngDNI.Find(What:=documento, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not celdaDNI Is Nothing Then
        nombreCompleto = wsMasterPersonal.Cells(celdaDNI.Row, wsMasterPersonal.ListObjects("PERSONAL").ListColumns("APELLIDOS Y NOMBRES").Index).Value
    Else
        MsgBox "Documento no encontrado en la tabla MASTER PERSONAL", vbExclamation
        Exit Sub
    End If
        
    ' Insertar una nueva fila en la tabla PersonalProyecto
    Set nuevaFila = tblPersonalProyecto.ListRows.Add
    
    ' Rellenar la nueva fila con los valores obtenidos
    With nuevaFila
        .Range(1, tblPersonalProyecto.ListColumns("DOC IDENTIDAD").Index).Value = "'" & documento
        .Range(1, tblPersonalProyecto.ListColumns("SERVICIO O PUESTO").Index).Value = ServicioOPuesto
        .Range(1, tblPersonalProyecto.ListColumns("EMPRESA").Index).Value = Empresa
        .Range(1, tblPersonalProyecto.ListColumns("PROYECTO").Index).Value = proyecto
        .Range(1, tblPersonalProyecto.ListColumns("NOMBRE COMPLETO").Index).Value = nombreCompleto
    End With
    
    MsgBox "Fila insertada correctamente en la tabla PersonalProyecto", vbInformation
End Sub


Sub InsertarVehiculoProyecto()
    Dim wsVehiculo As Worksheet
    Dim wsMasterVehiculo As Worksheet
    Dim tblVehiculoProyecto As ListObject
    Dim nuevaFila As ListRow
    Dim placa As String
    Dim Servicio As String
    Dim Empresa As String
    Dim proyecto As String
    Dim tipo As String
    Dim rngPlaca As Range
    Dim rngProyecto As Range
    Dim celdaExistente As Range
    Dim fila As Range
    Dim i As Long
    
    ' Definir las hojas
    Set wsVehiculo = ThisWorkbook.Sheets("VEHICULOS")
    Set wsMasterVehiculo = ThisWorkbook.Sheets("MASTER VEHICULOS")
    
    ' Obtener la tabla VehiculoProyecto
    On Error Resume Next
    Set tblVehiculoProyecto = wsVehiculo.ListObjects("VehiculoProyecto")
    On Error GoTo 0
    
    If tblVehiculoProyecto Is Nothing Then
        MsgBox "La tabla 'VehiculoProyecto' no existe en la hoja 'VEHICULOS'.", vbCritical
        Exit Sub
    End If
    
    ' Obtener los valores de las celdas
    placa = wsVehiculo.Range("H5").Value
    Servicio = wsVehiculo.Range("H4").Value
    Empresa = wsVehiculo.Range("H7").Value
    tipo = wsVehiculo.Range("H6").Value
    proyecto = wsVehiculo.Range("H2").Value
    
    ' Verificar si la tabla está vacía
    If tblVehiculoProyecto.ListRows.Count > 0 Then
        ' Buscar si la placa ya existe en la tabla VehiculoProyecto
        On Error Resume Next
        Set rngPlaca = tblVehiculoProyecto.ListColumns("PLACA").DataBodyRange
        Set rngProyecto = tblVehiculoProyecto.ListColumns("PROYECTO").DataBodyRange
        On Error GoTo 0
        
        ' Verificar si ambas columnas existen y tienen datos
        If Not rngPlaca Is Nothing And Not rngProyecto Is Nothing Then
            ' Recorrer las filas para buscar la coincidencia de PLACA + PROYECTO
            For i = 1 To rngPlaca.Rows.Count
                If rngPlaca.Cells(i, 1).Value = placa And rngProyecto.Cells(i, 1).Value = proyecto Then
                    ' Si encontramos la fila, la eliminamos
                    tblVehiculoProyecto.ListRows(i).Delete
                    Exit For
                End If
            Next i
        End If
    End If
    
    ' Si la placa ya existe, eliminar la fila existente
    If Not celdaExistente Is Nothing Then
        celdaExistente.EntireRow.Delete
    End If
    
    ' Insertar una nueva fila en la tabla VehiculoProyecto
    Set nuevaFila = tblVehiculoProyecto.ListRows.Add
    
    ' Rellenar la nueva fila con los valores obtenidos
    With nuevaFila
        .Range(1, tblVehiculoProyecto.ListColumns("PLACA").Index).Value = placa
        .Range(1, tblVehiculoProyecto.ListColumns("SERVICIO").Index).Value = Servicio
        .Range(1, tblVehiculoProyecto.ListColumns("TIPO DE VEHÍCULO").Index).Value = tipo
        .Range(1, tblVehiculoProyecto.ListColumns("EMPRESA").Index).Value = Empresa
        .Range(1, tblVehiculoProyecto.ListColumns("PROYECTO").Index).Value = proyecto
        .Range(1, tblVehiculoProyecto.ListColumns("ESTADO").Index).Value = "Activo"
    End With
    
    MsgBox "Fila insertada correctamente en la tabla Vehículo Proyecto", vbInformation
End Sub

