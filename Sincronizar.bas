Attribute VB_Name = "Sincronizar"
Sub SincronizarArchivos()

    ' Declaración de variables
    Dim libroOrigen As Workbook
    Dim libroDestino As Workbook
    Dim rutaOrigen As String
    Dim hojaDestino As Worksheet
    Dim hojaOrigen As Worksheet
    Dim tblDestino As ListObject
    Dim tblOrigen As ListObject
    
    ' Definir el libro destino (donde está la macro)
    Set libroDestino = ThisWorkbook
    
    
    ' Pregunta al usuario si desea continuar con el procedimiento
    confirmar = MsgBox("¿Desea continuar con el procedimiento?, va a borrar todo lo que tenga en su local, por lo que tenga el archivo central ", vbYesNo + vbQuestion, "Confirmación")
    If confirmar = vbNo Then
        Exit Sub ' Sale si el usuario selecciona No
    End If
    
    
    ' Obtener la ruta del libro origen desde la celda B8 de la hoja PRINCIPAL
    rutaOrigen = libroDestino.Sheets("PRINCIPAL").Range("B8").Value
    
    ' Verificar si la ruta del libro origen es válida
    If Dir(rutaOrigen) = "" Then
        MsgBox "El archivo de origen no se encuentra en la ruta especificada.", vbCritical
        Exit Sub
    End If
    
    ' Abrir el libro origen
    Set libroOrigen = Workbooks.Open(rutaOrigen)
    
    ' -------------------------
    ' Limpiar y copiar datos de las tablas de la hoja CATALOGOS
    ' -------------------------
    
    ' Proyectos
    LimpiarTabla libroDestino.Sheets("CATALOGOS").ListObjects("Proyectos")
    CopiarDatos libroOrigen.Sheets("CATALOGOS").ListObjects("Proyectos"), libroDestino.Sheets("CATALOGOS").ListObjects("Proyectos")
    
    ' Empresas
    LimpiarTabla libroDestino.Sheets("CATALOGOS").ListObjects("Empresas")
    CopiarDatos libroOrigen.Sheets("CATALOGOS").ListObjects("Empresas"), libroDestino.Sheets("CATALOGOS").ListObjects("Empresas")
    
    ' Puestos
    LimpiarTabla libroDestino.Sheets("CATALOGOS").ListObjects("Puestos")
    CopiarDatos libroOrigen.Sheets("CATALOGOS").ListObjects("Puestos"), libroDestino.Sheets("CATALOGOS").ListObjects("Puestos")
    
    ' GrupoPersonal
    LimpiarTabla libroDestino.Sheets("CATALOGOS").ListObjects("GrupoPersonal")
    CopiarDatos libroOrigen.Sheets("CATALOGOS").ListObjects("GrupoPersonal"), libroDestino.Sheets("CATALOGOS").ListObjects("GrupoPersonal")
    
    ' -------------------------
    ' Limpiar y copiar datos de la hoja MASTER PERSONAL
    ' -------------------------
    
    ' Tabla Personal
    LimpiarTabla libroDestino.Sheets("MASTER PERSONAL").ListObjects("Personal")
    CopiarDatos libroOrigen.Sheets("MASTER PERSONAL").ListObjects("Personal"), libroDestino.Sheets("MASTER PERSONAL").ListObjects("Personal")
    
    ' -------------------------
    ' Limpiar y copiar datos de la hoja MASTER VEHICULOS
    ' -------------------------
    
    ' Tabla Vehículo
    LimpiarTabla libroDestino.Sheets("MASTER VEHICULOS").ListObjects("Vehiculo")
    CopiarDatos libroOrigen.Sheets("MASTER VEHICULOS").ListObjects("Vehiculo"), libroDestino.Sheets("MASTER VEHICULOS").ListObjects("Vehiculo")
    
    ' -------------------------
    ' Limpiar y copiar datos de la hoja DOCUMENTO X PUESTO
    ' -------------------------
    
    ' Tabla DocumentoPorPuesto
    LimpiarTabla libroDestino.Sheets("DOCUMENTO X PUESTO").ListObjects("DocumentoPorPuesto")
    CopiarDatos libroOrigen.Sheets("DOCUMENTO X PUESTO").ListObjects("DocumentoPorPuesto"), libroDestino.Sheets("DOCUMENTO X PUESTO").ListObjects("DocumentoPorPuesto")
    
    ' -------------------------
    ' Limpiar y copiar datos de la hoja REQUISITOS X VEHICULO
    ' -------------------------
    
    ' Tabla RequisitosPorVehiculo
    LimpiarTabla libroDestino.Sheets("REQUISITOS X VEHICULO").ListObjects("RequisitosPorVehiculo")
    CopiarDatos libroOrigen.Sheets("REQUISITOS X VEHICULO").ListObjects("RequisitosPorVehiculo"), libroDestino.Sheets("REQUISITOS X VEHICULO").ListObjects("RequisitosPorVehiculo")
    
    ' -------------------------
    ' Limpiar y copiar datos de la hoja PERSONAL
    ' -------------------------
    
    ' Tabla PersonalProyecto
    LimpiarTabla libroDestino.Sheets("PERSONAL").ListObjects("PersonalProyecto")
    CopiarDatosPersonal libroOrigen.Sheets("PERSONAL").ListObjects("PersonalProyecto"), libroDestino.Sheets("PERSONAL").ListObjects("PersonalProyecto")
    
    ' -------------------------
    ' Limpiar y copiar datos de la hoja VEHICULOS
    ' -------------------------
    
    ' Tabla VehiculoProyecto
    LimpiarTabla libroDestino.Sheets("VEHICULOS").ListObjects("VehiculoProyecto")
    CopiarDatos libroOrigen.Sheets("VEHICULOS").ListObjects("VehiculoProyecto"), libroDestino.Sheets("VEHICULOS").ListObjects("VehiculoProyecto")
    
    ' -------------------------
    ' Limpiar y copiar datos de la hoja PERSONAL DOCUMENTOS (columnas A a E, filas desde la fila 2)
    ' -------------------------
    
    LimpiarRango libroDestino.Sheets("PERSONAL DOCUMENTOS").Range("A2:E" & libroDestino.Sheets("PERSONAL DOCUMENTOS").Cells(Rows.Count, "A").End(xlUp).Row)
    CopiarRango libroOrigen.Sheets("PERSONAL DOCUMENTOS").Range("A2:E" & libroOrigen.Sheets("PERSONAL DOCUMENTOS").Cells(Rows.Count, "A").End(xlUp).Row), _
        libroDestino.Sheets("PERSONAL DOCUMENTOS").Range("A2")
    
    ' -------------------------
    ' Limpiar y copiar datos de la hoja VEHICULO DOCUMENTOS (columnas A a E, filas desde la fila 2)
    ' -------------------------
    
    LimpiarRango libroDestino.Sheets("VEHICULO DOCUMENTOS").Range("A2:E" & libroDestino.Sheets("VEHICULO DOCUMENTOS").Cells(Rows.Count, "A").End(xlUp).Row)
    CopiarRango libroOrigen.Sheets("VEHICULO DOCUMENTOS").Range("A2:E" & libroOrigen.Sheets("VEHICULO DOCUMENTOS").Cells(Rows.Count, "A").End(xlUp).Row), _
        libroDestino.Sheets("VEHICULO DOCUMENTOS").Range("A2")


 ' -------------------------
    ' Limpiar y copiar datos de la hoja MASTER REQUISITOS VEHICULO (columnas A a F, filas desde la fila 2)
    ' -------------------------
    
    LimpiarRango libroDestino.Sheets("MASTER REQUISITOS VEHICULO").Range("A2:G" & libroDestino.Sheets("MASTER REQUISITOS VEHICULO").Cells(Rows.Count, "A").End(xlUp).Row)
    CopiarRango libroOrigen.Sheets("MASTER REQUISITOS VEHICULO").Range("A2:G" & libroOrigen.Sheets("MASTER REQUISITOS VEHICULO").Cells(Rows.Count, "A").End(xlUp).Row), _
        libroDestino.Sheets("MASTER REQUISITOS VEHICULO").Range("A2")


' -------------------------
    ' Limpiar y copiar datos de la hoja MASTER DOCUMENTOS PERSONAL (columnas A a F, filas desde la fila 2)
    ' -------------------------
    
    LimpiarRango libroDestino.Sheets("MASTER DOCUMENTOS PERSONAL").Range("A2:F" & libroDestino.Sheets("MASTER DOCUMENTOS PERSONAL").Cells(Rows.Count, "A").End(xlUp).Row)
    CopiarRango libroOrigen.Sheets("MASTER DOCUMENTOS PERSONAL").Range("A2:F" & libroOrigen.Sheets("MASTER DOCUMENTOS PERSONAL").Cells(Rows.Count, "A").End(xlUp).Row), _
        libroDestino.Sheets("MASTER DOCUMENTOS PERSONAL").Range("A2")


    ' Cerrar el libro origen
    libroOrigen.Close False
    
    Sheets("PRINCIPAL").Activate
    
    MsgBox "Sincronización completada", vbInformation

End Sub

' Subrutina para limpiar el contenido de una tabla
Sub LimpiarTabla(tbl As ListObject)
    On Error Resume Next
    tbl.DataBodyRange.Delete
    On Error GoTo 0
End Sub

' Subrutina para copiar datos entre dos tablas
Sub CopiarDatos(tblOrigen As ListObject, tblDestino As ListObject)
    ' Verificar que las tablas existan y tengan datos
    If Not tblOrigen Is Nothing And Not tblDestino Is Nothing Then
        If Not tblOrigen.DataBodyRange Is Nothing Then
            ' Verificar si la tabla de origen tiene datos
            If tblOrigen.DataBodyRange.Rows.Count > 0 Then
                ' Verificar si la tabla destino tiene datos antes de limpiar
                If Not tblDestino.DataBodyRange Is Nothing Then
                    ' Limpiar los datos actuales en la tabla destino
                    tblDestino.DataBodyRange.Delete
                End If
                
                ' Redimensionar la tabla destino para que coincida con la cantidad de filas y columnas de la tabla origen
                Dim numRows As Long
                Dim numCols As Long
                
                numRows = tblOrigen.DataBodyRange.Rows.Count
                numCols = tblOrigen.DataBodyRange.Columns.Count
                
                ' Asegurar que la tabla destino tenga suficiente espacio para los datos
                If tblDestino.ListRows.Count < numRows Then
                    ' Agregar filas si es necesario
                    tblDestino.Resize tblDestino.Range.Resize(numRows + 1, numCols)
                End If
                
                ' Copiar valores directamente sin usar Copy/PasteSpecial
                tblDestino.DataBodyRange.Cells(1, 1).Resize(numRows, numCols).Value = tblOrigen.DataBodyRange.Value
            Else
                MsgBox "La tabla de origen no contiene datos para copiar.", vbExclamation
            End If
        Else
            MsgBox "La tabla de origen está vacía o no existe.", vbExclamation
        End If
    Else
        MsgBox "No se encontró una de las tablas (origen o destino).", vbCritical
    End If
End Sub


Sub CopiarDatosPersonal(tblOrigen As ListObject, tblDestino As ListObject)
    ' Verificar que las tablas existan y tengan datos
    If Not tblOrigen Is Nothing And Not tblDestino Is Nothing Then
        If Not tblOrigen.DataBodyRange Is Nothing Then
            ' Verificar si la tabla de origen tiene datos
            If tblOrigen.DataBodyRange.Rows.Count > 0 Then
                ' Verificar si la tabla destino tiene datos antes de limpiar
                If Not tblDestino.DataBodyRange Is Nothing Then
                    ' Limpiar los datos actuales en la tabla destino
                    tblDestino.DataBodyRange.Delete
                End If
                
                ' Redimensionar la tabla destino para que coincida con la cantidad de filas y columnas de la tabla origen
                Dim numRows As Long
                Dim numCols As Long
                
                numRows = tblOrigen.DataBodyRange.Rows.Count
                numCols = tblOrigen.DataBodyRange.Columns.Count
                
                ' Asegurar que la tabla destino tenga suficiente espacio para los datos
                If tblDestino.ListRows.Count < numRows Then
                    ' Agregar filas si es necesario
                    tblDestino.Resize tblDestino.Range.Resize(numRows + 1, numCols)
                End If
                
                ' Copiar los datos de la primera columna como texto con un apóstrofo delante
                Dim i As Long
                For i = 1 To numRows
                    ' Añade un apóstrofo delante del valor de la primera columna
                    tblDestino.DataBodyRange.Cells(i, 1).Value = "'" & tblOrigen.DataBodyRange.Cells(i, 1).Value
                Next i
                
                ' Copiar el resto de las columnas sin cambios
                If numCols > 1 Then
                    tblDestino.DataBodyRange.Cells(1, 2).Resize(numRows, numCols - 1).Value = tblOrigen.DataBodyRange.Cells(1, 2).Resize(numRows, numCols - 1).Value
                End If
            Else
                MsgBox "La tabla de origen no contiene datos para copiar.", vbExclamation
            End If
        Else
            MsgBox "La tabla de origen está vacía o no existe.", vbExclamation
        End If
    Else
        MsgBox "No se encontró una de las tablas (origen o destino).", vbCritical
    End If
End Sub



' Subrutina para limpiar el contenido de un rango
Sub LimpiarRango(rng As Range)
    rng.ClearContents
End Sub

' Subrutina para copiar un rango de celdas a otro
Sub CopiarRango(rngOrigen As Range, rngDestino As Range)
    rngOrigen.Copy
    rngDestino.PasteSpecial xlPasteValues
End Sub



Sub SincronizarArchivoCentral()

    ' Declaración de variables
    Dim libroOrigen As Workbook
    Dim libroDestino As Workbook
    Dim rutaOrigen As String
    Dim hojaDestino As Worksheet
    Dim hojaOrigen As Worksheet
    Dim tblDestino As ListObject
    Dim tblOrigen As ListObject
    
    ' Definir el libro destino (donde está la macro)
    Set libroDestino = ThisWorkbook
    
    
    ' Pregunta al usuario si desea continuar con el procedimiento lo
    confirmar = MsgBox("¿Desea continuar con el procedimiento?, esto copiará todo lo que hizo localmente en el archivo central", vbYesNo + vbQuestion, "Confirmación")
    If confirmar = vbNo Then
        Exit Sub ' Sale si el usuario selecciona No
    End If
    
    
    ' Obtener la ruta del libro origen desde la celda B8 de la hoja PRINCIPAL
    rutaOrigen = libroDestino.Sheets("PRINCIPAL").Range("B8").Value
    
    ' Verificar si la ruta del libro origen es válida
    If Dir(rutaOrigen) = "" Then
        MsgBox "El archivo de origen no se encuentra en la ruta especificada.", vbCritical
        Exit Sub
    End If
    
    ' Abrir el libro origen
    Set libroOrigen = Workbooks.Open(rutaOrigen)
    
    ' -------------------------
    ' Limpiar y copiar datos de las tablas de la hoja CATALOGOS
    ' -------------------------
    
    ' Proyectos
    LimpiarTabla libroOrigen.Sheets("CATALOGOS").ListObjects("Proyectos")
    CopiarDatos libroDestino.Sheets("CATALOGOS").ListObjects("Proyectos"), libroOrigen.Sheets("CATALOGOS").ListObjects("Proyectos")
    
    
    ' Empresas
    LimpiarTabla libroOrigen.Sheets("CATALOGOS").ListObjects("Empresas")
    CopiarDatos libroDestino.Sheets("CATALOGOS").ListObjects("Empresas"), libroOrigen.Sheets("CATALOGOS").ListObjects("Empresas")
    
    ' Puestos
    LimpiarTabla libroOrigen.Sheets("CATALOGOS").ListObjects("Puestos")
    CopiarDatos libroDestino.Sheets("CATALOGOS").ListObjects("Puestos"), libroOrigen.Sheets("CATALOGOS").ListObjects("Puestos")
    
    ' GrupoPersonal
    LimpiarTabla libroOrigen.Sheets("CATALOGOS").ListObjects("GrupoPersonal")
    CopiarDatos libroDestino.Sheets("CATALOGOS").ListObjects("GrupoPersonal"), libroOrigen.Sheets("CATALOGOS").ListObjects("GrupoPersonal")
    
    ' -------------------------
    ' Limpiar y copiar datos de la hoja MASTER PERSONAL
    ' -------------------------
    
    ' Tabla Personal
    LimpiarTabla libroOrigen.Sheets("MASTER PERSONAL").ListObjects("Personal")
    CopiarDatos libroDestino.Sheets("MASTER PERSONAL").ListObjects("Personal"), libroOrigen.Sheets("MASTER PERSONAL").ListObjects("Personal")
    
    ' -------------------------
    ' Limpiar y copiar datos de la hoja MASTER VEHICULOS
    ' -------------------------
    
    ' Tabla Vehículo
    LimpiarTabla libroOrigen.Sheets("MASTER VEHICULOS").ListObjects("Vehiculo")
    CopiarDatos libroDestino.Sheets("MASTER VEHICULOS").ListObjects("Vehiculo"), libroOrigen.Sheets("MASTER VEHICULOS").ListObjects("Vehiculo")
    
    ' -------------------------
    ' Limpiar y copiar datos de la hoja DOCUMENTO X PUESTO
    ' -------------------------
    
    ' Tabla DocumentoPorPuesto
    LimpiarTabla libroOrigen.Sheets("DOCUMENTO X PUESTO").ListObjects("DocumentoPorPuesto")
    CopiarDatos libroDestino.Sheets("DOCUMENTO X PUESTO").ListObjects("DocumentoPorPuesto"), libroOrigen.Sheets("DOCUMENTO X PUESTO").ListObjects("DocumentoPorPuesto")
    
    ' -------------------------
    ' Limpiar y copiar datos de la hoja REQUISITOS X VEHICULO
    ' -------------------------
    
    ' Tabla RequisitosPorVehiculo
    LimpiarTabla libroOrigen.Sheets("REQUISITOS X VEHICULO").ListObjects("RequisitosPorVehiculo")
    CopiarDatos libroDestino.Sheets("REQUISITOS X VEHICULO").ListObjects("RequisitosPorVehiculo"), libroOrigen.Sheets("REQUISITOS X VEHICULO").ListObjects("RequisitosPorVehiculo")
    
    ' -------------------------
    ' Limpiar y copiar datos de la hoja PERSONAL
    ' -------------------------
    
    ' Tabla PersonalProyecto
    LimpiarTabla libroOrigen.Sheets("PERSONAL").ListObjects("PersonalProyecto")
    CopiarDatosPersonal libroDestino.Sheets("PERSONAL").ListObjects("PersonalProyecto"), libroOrigen.Sheets("PERSONAL").ListObjects("PersonalProyecto")
    
    ' -------------------------
    ' Limpiar y copiar datos de la hoja VEHICULOS
    ' -------------------------
    
    ' Tabla VehiculoProyecto
    LimpiarTabla libroOrigen.Sheets("VEHICULOS").ListObjects("VehiculoProyecto")
    CopiarDatos libroDestino.Sheets("VEHICULOS").ListObjects("VehiculoProyecto"), libroOrigen.Sheets("VEHICULOS").ListObjects("VehiculoProyecto")
    
    ' -------------------------
    ' Limpiar y copiar datos de la hoja PERSONAL DOCUMENTOS (columnas A a E, filas desde la fila 2)
    ' -------------------------
    
    LimpiarRango libroOrigen.Sheets("PERSONAL DOCUMENTOS").Range("A2:E" & libroOrigen.Sheets("PERSONAL DOCUMENTOS").Cells(Rows.Count, "A").End(xlUp).Row)
    CopiarRango libroDestino.Sheets("PERSONAL DOCUMENTOS").Range("A2:E" & libroDestino.Sheets("PERSONAL DOCUMENTOS").Cells(Rows.Count, "A").End(xlUp).Row), _
        libroOrigen.Sheets("PERSONAL DOCUMENTOS").Range("A2")
    
    ' -------------------------
    ' Limpiar y copiar datos de la hoja VEHICULO DOCUMENTOS (columnas A a E, filas desde la fila 2)
    ' -------------------------
    
    LimpiarRango libroOrigen.Sheets("VEHICULO DOCUMENTOS").Range("A2:E" & libroOrigen.Sheets("VEHICULO DOCUMENTOS").Cells(Rows.Count, "A").End(xlUp).Row)
    CopiarRango libroDestino.Sheets("VEHICULO DOCUMENTOS").Range("A2:E" & libroDestino.Sheets("VEHICULO DOCUMENTOS").Cells(Rows.Count, "A").End(xlUp).Row), _
        libroOrigen.Sheets("VEHICULO DOCUMENTOS").Range("A2")


 ' -------------------------
    ' Limpiar y copiar datos de la hoja MASTER REQUISITOS VEHICULO (columnas A a F, filas desde la fila 2)
    ' -------------------------
    
    LimpiarRango libroOrigen.Sheets("MASTER REQUISITOS VEHICULO").Range("A2:G" & libroOrigen.Sheets("MASTER REQUISITOS VEHICULO").Cells(Rows.Count, "A").End(xlUp).Row)
    CopiarRango libroDestino.Sheets("MASTER REQUISITOS VEHICULO").Range("A2:G" & libroDestino.Sheets("MASTER REQUISITOS VEHICULO").Cells(Rows.Count, "A").End(xlUp).Row), _
        libroOrigen.Sheets("MASTER REQUISITOS VEHICULO").Range("A2")


' -------------------------
    ' Limpiar y copiar datos de la hoja MASTER DOCUMENTOS PERSONAL (columnas A a F, filas desde la fila 2)
    ' -------------------------
    
    LimpiarRango libroOrigen.Sheets("MASTER DOCUMENTOS PERSONAL").Range("A2:F" & libroOrigen.Sheets("MASTER DOCUMENTOS PERSONAL").Cells(Rows.Count, "A").End(xlUp).Row)
    CopiarRango libroDestino.Sheets("MASTER DOCUMENTOS PERSONAL").Range("A2:F" & libroDestino.Sheets("MASTER DOCUMENTOS PERSONAL").Cells(Rows.Count, "A").End(xlUp).Row), _
        libroOrigen.Sheets("MASTER DOCUMENTOS PERSONAL").Range("A2")


    ' Cerrar el libro origen
    libroOrigen.Close True
    
    Sheets("PRINCIPAL").Activate
    
    MsgBox "Se actualizó información a archivo central", vbInformation

End Sub


