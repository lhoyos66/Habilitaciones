Attribute VB_Name = "Utilitarios"
Function UniqueValuesVertical(rng As Range) As Variant
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim cell As Range
    For Each cell In rng
        If Not dict.Exists(cell.Value) Then
            dict.Add cell.Value, Nothing
        End If
    Next cell
    
    ' Convert dictionary keys to vertical array
    Dim result() As Variant
    ReDim result(1 To dict.Count, 1 To 1)
    
    Dim i As Integer
    i = 1
    Dim key As Variant
    For Each key In dict.keys
        result(i, 1) = key
        i = i + 1
    Next key
    
    UniqueValuesVertical = result
End Function

Function UniqueNonEmptyValues(rng As Range) As Variant
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim cell As Range
    For Each cell In rng
        If Not IsEmpty(cell.Value) And Not dict.Exists(cell.Value) Then
            dict.Add cell.Value, Nothing
        End If
    Next cell
    
    ' Convert dictionary keys to vertical array
    Dim result() As Variant
    ReDim result(1 To dict.Count, 1 To 1)
    
    Dim i As Integer
    i = 1
    Dim key As Variant
    For Each key In dict.keys
        result(i, 1) = key
        i = i + 1
    Next key
    
    UniqueNonEmptyValues = result
End Function

Function FiltrarUnicosTabla(tabla As ListObject, colCriterio As String, valorCriterio As Variant, colResultados As String) As Variant
    Dim i As Long
    Dim datos As Variant
    Dim resultadosTemp() As Variant
    Dim dict As Object
    Dim colCritIndex As Long
    Dim colResIndex As Long
    Dim filas As Long

    ' Crear un diccionario para almacenar valores únicos
    Set dict = CreateObject("Scripting.Dictionary")

    ' Obtener los datos del rango en una matriz
    On Error GoTo ErrorHandler
    datos = tabla.DataBodyRange.Value
    Debug.Print "Datos obtenidos"

    ' Contar cuántas filas tiene el rango
    filas = UBound(datos, 1)
    Debug.Print "Número de filas: " & filas

    ' Encontrar el índice de la columna de criterio y de resultados
    colCritIndex = Application.Match(colCriterio, Application.Index(tabla.HeaderRowRange.Value, 1, 0), 0)
    colResIndex = Application.Match(colResultados, Application.Index(tabla.HeaderRowRange.Value, 1, 0), 0)
    
    If IsError(colCritIndex) Or IsError(colResIndex) Then
        Debug.Print "Error en los índices de las columnas"
        FiltrarUnicosTabla = CVErr(xlErrRef)
        Exit Function
    End If

    Debug.Print "Índice de columna de criterio: " & colCritIndex
    Debug.Print "Índice de columna de resultados: " & colResIndex

    ' Recorrer cada fila del rango de datos
    For i = 1 To filas
        ' Comprobar si la fila cumple el criterio
        If datos(i, colCritIndex) = valorCriterio Then
            Debug.Print "Criterio cumplido en fila: " & i
            ' Agregar el valor único al diccionario
            If Not dict.Exists(datos(i, colResIndex)) Then
                dict.Add datos(i, colResIndex), Nothing
            End If
        End If
    Next i

    ' Crear una matriz para almacenar los resultados únicos
    If dict.Count > 0 Then
        ReDim resultadosTemp(1 To dict.Count, 1 To 1)
        i = 1
        For Each key In dict.keys
            resultadosTemp(i, 1) = key
            i = i + 1
        Next key
        FiltrarUnicosTabla = resultadosTemp
    Else
        FiltrarUnicosTabla = CVErr(xlErrNA) ' Devolver #N/A si no hay coincidencias
    End If

    Exit Function
    
ErrorHandler:
    Debug.Print "Error: " & Err.Description
    FiltrarUnicosTabla = CVErr(xlErrValue)
End Function

Sub ActivarPantalla()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ActualizarPersonalDocumentos ("PERSONAL DOCUMENTOS HABIL")
    ActualizarPersonalDocumentos ("PERSONAL DOCUMENTOS RRHH")
    MsgBox "Los datos se han actualizado correctamente en PERSONAL DOCUMENTOS.", vbInformation
End Sub


Sub BackupArchivosPersonal()
    Dim ws As Worksheet
    Dim wsPrincipal As Worksheet
    Dim tbl As ListObject
    Dim fd As FileDialog
    Dim destFolder As String
    Dim projectFolder As String
    Dim personalFolder As String
    Dim companyFolder As String
    Dim docFolder As String
    Dim projectCode As String
    Dim companyName As String
    Dim docID As String
    Dim filePath As String
    Dim fullFilePath As String
    Dim carpetaBase As String
    Dim lastRow As Long
    Dim i As Long
    Dim fso As Object
    Dim currentPath As String
    Dim destino As String

    ' Crear objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Obtener la ruta del directorio actual
    currentPath = fso.GetAbsolutePathName(".")
    Debug.Print "Directorio actual: " & currentPath

    ' Configurar la hoja de trabajo y la tabla
    Set ws = ThisWorkbook.Sheets("DASHBOARD PERSONAL")
    Set wsPrincipal = ThisWorkbook.Sheets("PRINCIPAL")
    Set tbl = ws.ListObjects("DashboardPersonal")
    
    ' Crear ventana de diálogo para seleccionar la carpeta de destino
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Seleccionar carpeta de destino para el backup"
        If .Show = -1 Then
            destFolder = .SelectedItems(1)
        Else
            MsgBox "No se seleccionó ninguna carpeta. Proceso cancelado.", vbExclamation
            Exit Sub
        End If
    End With
    
    Debug.Print "Carpeta de destino seleccionada: " & destFolder

    carpetaBase = wsPrincipal.Cells(6, 2).Value
    Debug.Print "Carpeta base: " & carpetaBase
    
    ' Recorrer cada fila de la tabla
    lastRow = tbl.ListRows.Count
    For i = 1 To lastRow
        ' Leer valores de las columnas
        projectCode = tbl.ListColumns("CÓDIGO PROYECTO").DataBodyRange(i, 1).Value
        companyName = tbl.ListColumns("EMPRESA").DataBodyRange(i, 1).Value
        docID = tbl.ListColumns("DOC IDENTIDAD").DataBodyRange(i, 1).Value
        filePath = tbl.ListColumns("LINK").DataBodyRange(i, 1).Hyperlinks(1).Address
        
        Debug.Print "Procesando fila: " & i
        Debug.Print "Código de proyecto: " & projectCode
        Debug.Print "Nombre de la empresa: " & companyName
        Debug.Print "Documento de identidad: " & docID
        Debug.Print "Ruta del archivo: " & filePath
        
        ' Obtener la ruta completa si es relativa
        'If Left(filePath, 2) = ".." Or Left(filePath, 1) = "\" Then
            'fullFilePath = fso.GetAbsolutePathName(fso.BuildPath(currentPath, filePath))
        'Else
            fullFilePath = filePath
        'End If
        
        Debug.Print "Ruta completa del archivo: " & fullFilePath

        ' Crear rutas de las carpetas
        projectFolder = destFolder & "\" & projectCode
        personalFolder = projectFolder & "\PERSONAL"
        companyFolder = personalFolder & "\" & companyName
        docFolder = companyFolder & "\" & docID
        
        ' Crear las carpetas si no existen
        CrearCarpetas projectFolder
        CrearCarpetas personalFolder
        CrearCarpetas companyFolder
        CrearCarpetas docFolder
        
        destino = docFolder & "\" & fso.GetFileName(fullFilePath)

        CrearCarpetas destino
        ' Verificar la ruta de destino
        Debug.Print "Ruta de destino del archivo: " & destino
        
        Set folder = fso.GetFolder(fullFilePath)
        If folder.Files.Count > 0 Then
            ' Copiar todos los archivos de la carpeta de origen a la carpeta de destino
                For Each file In folder.Files
                    Debug.Print "Copiando archivo desde: " & file.Path & " a " & destino & "\" & file.Name
                    fso.CopyFile Source:=file.Path, Destination:=destino & "\" & file.Name
                    Debug.Print "Archivo copiado: " & destino & "\" & file.Name
                Next file
        Else
            Debug.Print "Archivo no encontrado: " & fullFilePath
        End If
    Next i
    
    MsgBox "Backup completado exitosamente.", vbInformation
End Sub

Sub CrearCarpetas(ruta As String)
    Dim fso As Object
    Dim parts As Variant
    Dim currentPath As String
    Dim i As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    parts = Split(ruta, "\")
    
    currentPath = parts(0)
    For i = 1 To UBound(parts)
        currentPath = currentPath & "\" & parts(i)
        If Not fso.FolderExists(currentPath) Then
            fso.CreateFolder (currentPath)
            Debug.Print "Carpeta creada: " & currentPath
        End If
    Next i
End Sub

Sub BackupArchivosVehiculo()
    Dim ws As Worksheet
    Dim wsPrincipal As Worksheet
    Dim tbl As ListObject
    Dim fd As FileDialog
    Dim destFolder As String
    Dim projectFolder As String
    Dim vehiculoFolder As String
    Dim companyFolder As String
    Dim docFolder As String
    Dim projectCode As String
    Dim companyName As String
    Dim docID As String
    Dim filePath As String
    Dim fullFilePath As String
    Dim carpetaBase As String
    Dim lastRow As Long
    Dim i As Long
    Dim fso As Object
    Dim currentPath As String
    Dim destino As String

    ' Crear objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Obtener la ruta del directorio actual
    currentPath = fso.GetAbsolutePathName(".")
    Debug.Print "Directorio actual: " & currentPath

    ' Configurar la hoja de trabajo y la tabla
    Set ws = ThisWorkbook.Sheets("DASHBOARD VEHICULO")
    Set wsPrincipal = ThisWorkbook.Sheets("PRINCIPAL")
    Set tbl = ws.ListObjects("DashboardVehiculo")
    
    ' Crear ventana de diálogo para seleccionar la carpeta de destino
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Seleccionar carpeta de destino para el backup"
        If .Show = -1 Then
            destFolder = .SelectedItems(1)
        Else
            MsgBox "No se seleccionó ninguna carpeta. Proceso cancelado.", vbExclamation
            Exit Sub
        End If
    End With
    
    Debug.Print "Carpeta de destino seleccionada: " & destFolder

    carpetaBase = wsPrincipal.Cells(6, 2).Value
    Debug.Print "Carpeta base: " & carpetaBase
    
    ' Recorrer cada fila de la tabla
    lastRow = tbl.ListRows.Count
    For i = 1 To lastRow
        ' Leer valores de las columnas
        projectCode = tbl.ListColumns("CÓDIGO PROYECTO").DataBodyRange(i, 1).Value
        companyName = tbl.ListColumns("EMPRESA").DataBodyRange(i, 1).Value
        docID = tbl.ListColumns("PLACA").DataBodyRange(i, 1).Value
        filePath = tbl.ListColumns("LINK").DataBodyRange(i, 1).Hyperlinks(1).Address
        
        Debug.Print "Procesando fila: " & i
        Debug.Print "Código de proyecto: " & projectCode
        Debug.Print "Nombre de la empresa: " & companyName
        Debug.Print "Placa Vehículo: " & docID
        Debug.Print "Ruta del archivo: " & filePath
        
        ' Obtener la ruta completa si es relativa
        If Left(filePath, 2) = ".." Or Left(filePath, 1) = "\" Then
            fullFilePath = fso.GetAbsolutePathName(fso.BuildPath(currentPath, filePath))
        Else
            fullFilePath = filePath
        End If
        
        Debug.Print "Ruta completa del archivo: " & fullFilePath

        ' Crear rutas de las carpetas
        projectFolder = destFolder & "\" & projectCode
        vehiculoFolder = projectFolder & "\VEHICULOS"
        companyFolder = vehiculoFolder & "\" & companyName
        docFolder = companyFolder & "\" & docID
        
        ' Crear las carpetas si no existen
        CrearCarpetas projectFolder
        CrearCarpetas vehiculoFolder
        CrearCarpetas companyFolder
        CrearCarpetas docFolder
        
        destino = docFolder & "\" & fso.GetFileName(fullFilePath)

        CrearCarpetas destino
        ' Verificar la ruta de destino
        Debug.Print "Ruta de destino del archivo: " & destino
        
        Set folder = fso.GetFolder(fullFilePath)
        If folder.Files.Count > 0 Then
            ' Copiar todos los archivos de la carpeta de origen a la carpeta de destino
                For Each file In folder.Files
                    Debug.Print "Copiando archivo desde: " & file.Path & " a " & destino & "\" & file.Name
                    fso.CopyFile Source:=file.Path, Destination:=destino & "\" & file.Name
                    Debug.Print "Archivo copiado: " & destino & "\" & file.Name
                Next file
        Else
            Debug.Print "Archivo no encontrado: " & fullFilePath
        End If
    Next i
    
    MsgBox "Backup completado exitosamente.", vbInformation
End Sub



Sub ResumenHabilitacionesPorTrabajador()
    Dim wsDocumentoPuesto As Worksheet
    Dim wsPersonal As Worksheet
    Dim wsPersonalDocumentos As Worksheet
    Dim wsGestionIntegral As Worksheet
    Dim wsDocXPuesto As Worksheet 'Modificación
    Dim wsPrincipal As Worksheet
    Dim wsDocPorPuesto As Worksheet
    Dim docPorPuestoDict As Object
    Dim docXPuestoDictEmp As Object 'Modificación
    Dim dictDocumentos As Object
    Dim dictTrabajadores As Object
    Dim celda As Range
    Dim lastRow As Long
    Dim lastRow1 As ListRow
    Dim proyectoCodigo As String
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim trabajador As Variant
    Dim documento As Variant
    Dim docPorPuestoData As Variant
    Dim docXPuestoData As Variant
    Dim DocIdentidad As String
    Dim nombre As String
    Dim cargoPuesto As String
    Dim encontrado As Boolean
    Dim tbl As ListObject
    Dim headerRange As Range
    Dim keyObligatorio As String
    Dim docObligatorio As String
    Dim EstadoDoc As String
   

    ' Inicializar diccionarios
    Set dictDocumentos = CreateObject("Scripting.Dictionary")
    Set dictTrabajadores = CreateObject("Scripting.Dictionary")
    Set docPorPuestoDict = CreateObject("Scripting.Dictionary")
    Set docXPuestoDictEmp = CreateObject("Scripting.Dictionary") 'Modificación

    ' Establecer referencias a las hojas
    Set wsDocumentoPuesto = ThisWorkbook.Sheets("DOCUMENTO X PUESTO")
    Set wsPersonal = ThisWorkbook.Sheets("PERSONAL")
    Set wsPersonalDocumentos = ThisWorkbook.Sheets("PERSONAL DOCUMENTOS")
    Set wsGestionIntegral = ThisWorkbook.Sheets("GESTION INTEGRAL CAMPO")
    Set wsPrincipal = ThisWorkbook.Sheets("PRINCIPAL")
    Set wsDocXPuesto = ThisWorkbook.Sheets("MASTER DOCUMENTOS PERSONAL") 'Modificación

    'Modificación
    docXPuestoData = wsDocXPuesto.Range("A2:D" & wsDocXPuesto.Cells(wsDocXPuesto.Rows.Count, "A").End(xlUp).Row).Value
    For i = LBound(docXPuestoData, 1) To UBound(docXPuestoData, 1)
        docXPuestoDictEmp(docXPuestoData(i, 2)) = docXPuestoData(i, 4)
    Next i

    docPorPuestoData = wsDocumentoPuesto.Range("A2:E" & wsDocumentoPuesto.Cells(wsDocumentoPuesto.Rows.Count, "A").End(xlUp).Row).Value
    For i = LBound(docPorPuestoData, 1) To UBound(docPorPuestoData, 1)
        docPorPuestoDict(docPorPuestoData(i, 4) & "-" & docPorPuestoData(i, 1) & "-" & docPorPuestoData(i, 2)) = docPorPuestoData(i, 3)
    Next i

    ' Obtener el código del proyecto de la celda A2 en la hoja PRINCIPAL
    proyectoCodigo = wsPrincipal.Range("A2").Value
    

    ' Obtener documentos únicos de la tabla DocumentoPorPuesto
    lastRow = wsDocumentoPuesto.Cells(wsDocumentoPuesto.Rows.Count, "B").End(xlUp).Row
    For Each celda In wsDocumentoPuesto.Range("B2:B" & lastRow)
        If celda.Offset(0, 2).Value = proyectoCodigo Then
            documento = celda.Value
            If Not dictDocumentos.Exists(documento) Then
                dictDocumentos.Add documento, documento
            End If
        End If
    Next celda

    ' Obtener trabajadores del proyecto de la tabla PersonalProyecto
    lastRow = wsPersonal.Cells(wsPersonal.Rows.Count, "A").End(xlUp).Row
    For Each celda In wsPersonal.Range("A2:A" & lastRow)
        If celda.Offset(0, 4).Value = proyectoCodigo Then
            DocIdentidad = celda.Value
            trabajador = celda.Offset(0, 1).Value
            If Not dictTrabajadores.Exists(trabajador) Then
                dictTrabajadores.Add trabajador, Array(trabajador, celda.Offset(0, 5).Value, celda.Offset(0, 2).Value, DocIdentidad, celda.Offset(0, 6).Value)
            End If
        End If
    Next celda

    ' Limpiar la hoja de destino
    wsGestionIntegral.Cells.Clear

    ' Crear encabezados de la tabla en la hoja GESTION INTEGRAL CAMPO
    wsGestionIntegral.Range("A1").Value = "NOMBRE"
    wsGestionIntegral.Range("B1").Value = "GRUPO"
    wsGestionIntegral.Range("C1").Value = "CARGO"
    wsGestionIntegral.Range("D1").Value = "ESTADO"
    colIndex = 5
    'Modificación
    For Each documento In dictDocumentos.keys
        protocolo = UCase(docXPuestoDictEmp(documento))
        
        ' Escribe el valor del documento en la celda
        wsGestionIntegral.Cells(1, colIndex).Value = documento
        
        ' Verifica si el protocolo es distinto de "TEMA"
        If protocolo <> "TEMA" Then
            ' Si el protocolo es diferente de "TEMA", el fondo es naranja
            wsGestionIntegral.Cells(1, colIndex).Interior.Color = RGB(0, 180, 0) ' Verde
        Else
            ' Si el protocolo es igual a "TEMA", el fondo es azul
           wsGestionIntegral.Cells(1, colIndex).Interior.Color = RGB(21, 96, 130) ' Azul personalizado de la imagen
       End If
        
        ' Incrementa el índice de la columna
        colIndex = colIndex + 1
    Next documento



    ' Llenar la tabla con los trabajadores y sus documentos
    rowIndex = 2
    Dim observacion As String
    For Each trabajador In dictTrabajadores.keys
        wsGestionIntegral.Cells(rowIndex, 1).Value = dictTrabajadores(trabajador)(0)
        wsGestionIntegral.Cells(rowIndex, 2).Value = dictTrabajadores(trabajador)(1)
        wsGestionIntegral.Cells(rowIndex, 3).Value = dictTrabajadores(trabajador)(2)
        wsGestionIntegral.Cells(rowIndex, 4).Value = dictTrabajadores(trabajador)(4)
        carpoPuesto = dictTrabajadores(trabajador)(2)
        DocIdentidad = dictTrabajadores(trabajador)(3)
        Nombres = dictTrabajadores(trabajador)(0)
        
        ' Verificar documentos presentados
        colIndex = 5
        For Each documento In dictDocumentos.keys
            encontrado = False
            encontrado1 = False
            observacion = ""
            lastRow = wsPersonalDocumentos.Cells(wsPersonalDocumentos.Rows.Count, "A").End(xlUp).Row
            For Each celda In wsPersonalDocumentos.Range("A2:A" & lastRow)
                keyObligatorio = proyectoCodigo & "-" & carpoPuesto & "-" & documento
                docObligatorio = docPorPuestoDict(keyObligatorio)
                If celda.Offset(0, 1).Value = documento And celda.Offset(0, 0).Value = DocIdentidad Then
                    observacion = celda.Offset(0, 4).Value 'Modificación
                End If
                If docObligatorio = "SI" Then
                    EstadoDoc = ""
                    If celda.Value = DocIdentidad And celda.Offset(0, 1).Value = documento Then
                        If celda.Offset(0, 3).Value < Date Then
                            EstadoDoc = "VENCIDO"
                        End If
                        encontrado = True
                        Exit For
                    End If
                Else
                    docObligatorio = "NO"
                    EstadoDoc = ""
                    If celda.Value = DocIdentidad And celda.Offset(0, 1).Value = documento Then
                        If celda.Offset(0, 3).Value < Date Then
                            EstadoDoc = "VENCIDO"
                        Else
                            EstadoDoc = "VÁLIDO"
                        End If
                        encontrado = True
                        Exit For
                    End If
                End If
            Next celda
            If encontrado And docObligatorio = "SI" Then
                 If EstadoDoc = "" And observacion <> "APROBADO POR GERENCIA" And observacion <> "NO RESTRICTIVO" And observacion <> "NO APLICA" And observacion <> "OBSERVADO" Then
                    wsGestionIntegral.Cells(rowIndex, colIndex).Value = "X"
                 Else
                    If observacion = "APROBADO POR GERENCIA" Then
                        wsGestionIntegral.Cells(rowIndex, colIndex).Value = "AGA"
                    Else
                        If observacion = "NO RESTRICTIVO" Then
                            wsGestionIntegral.Cells(rowIndex, colIndex).Value = "NR"
                        Else
                            If observacion = "NO APLICA" Then
                                wsGestionIntegral.Cells(rowIndex, colIndex).Value = "NA"
                            Else
                                If observacion = "OBSERVADO" Then
                                    wsGestionIntegral.Cells(rowIndex, colIndex).Value = "OBSERVADO"
                                Else
                                    wsGestionIntegral.Cells(rowIndex, colIndex).Value = "VENCIDO"
                                End If
                            End If
                        End If
                    End If
                 End If
            Else
                If docObligatorio = "NO" Then
                    If EstadoDoc = "" Or EstadoDoc = "VENCIDO" Then
                        wsGestionIntegral.Cells(rowIndex, colIndex).Value = "NA"
                    Else
                        If observacion = "APROBADO POR GERENCIA" Then
                            wsGestionIntegral.Cells(rowIndex, colIndex).Value = "AGA"
                        Else
                            If observacion = "NO RESTRICTIVO" Then
                                wsGestionIntegral.Cells(rowIndex, colIndex).Value = "NR"
                            Else
                                If observacion = "NO APLICA" Then
                                    wsGestionIntegral.Cells(rowIndex, colIndex).Value = "NA"
                                Else
                                    If observacion = "OBSERVADO" Then
                                        wsGestionIntegral.Cells(rowIndex, colIndex).Value = "OBSERVADO"
                                    Else
                                        wsGestionIntegral.Cells(rowIndex, colIndex).Value = "X"
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    If (encontrado = False Or encontrado = Falso) And observacion <> "APROBADO POR GERENCIA" And observacion <> "NO RESTRICTIVO" And observacion <> "NO APLICA" Then
                        wsGestionIntegral.Cells(rowIndex, colIndex).Value = "FALTA"
                    Else
                    'Modificación
                        If observacion = "APROBADO POR GERENCIA" Then
                            wsGestionIntegral.Cells(rowIndex, colIndex).Value = "AGA"
                        Else
                            If observacion = "NO RESTRICTIVO" Then
                                wsGestionIntegral.Cells(rowIndex, colIndex).Value = "NR"
                            Else
                                If observacion = "NO APLICA" Then
                                    wsGestionIntegral.Cells(rowIndex, colIndex).Value = "NA"
                                Else
                                    wsGestionIntegral.Cells(rowIndex, colIndex).Value = "FALTA"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            ' Aplica formato condicional si el valor de la celda es "FALTA"
            If wsGestionIntegral.Cells(rowIndex, colIndex).Value = "FALTA" Then
                With wsGestionIntegral.Cells(rowIndex, colIndex)
                    .Interior.Color = RGB(255, 0, 0) ' Fondo rojo
                    .Font.Color = RGB(255, 255, 255) ' Letras blancas
                End With
            End If
            ' Aplica formato condicional si el valor de la celda es "VENCIDO"
            If wsGestionIntegral.Cells(rowIndex, colIndex).Value = "VENCIDO" Then
                With wsGestionIntegral.Cells(rowIndex, colIndex)
                    .Interior.Color = RGB(255, 165, 0) ' Fondo rojo
                    .Font.Color = RGB(255, 255, 255) ' Letras blancas
                End With
            End If
            
            If wsGestionIntegral.Cells(rowIndex, colIndex).Value = "OBSERVADO" Then
                With wsGestionIntegral.Cells(rowIndex, colIndex)
                    .Interior.Color = RGB(128, 0, 128) ' Fondo rojo
                    .Font.Color = RGB(255, 255, 255) ' Letras blancas
                End With
            End If

            colIndex = colIndex + 1
        Next documento
        
        rowIndex = rowIndex + 1
    Next trabajador

    ' Eliminar la última fila si está vacía
    If Application.WorksheetFunction.CountA(wsGestionIntegral.Rows(rowIndex - 1)) = 0 Then
        wsGestionIntegral.Rows(rowIndex - 1).Delete
        rowIndex = rowIndex - 1
    End If

    ' Crear tabla
    Set headerRange = wsGestionIntegral.Range(wsGestionIntegral.Cells(1, 1), wsGestionIntegral.Cells(1, dictDocumentos.Count + 4))
    Set tbl = wsGestionIntegral.ListObjects.Add(xlSrcRange, headerRange.Resize(rowIndex, dictDocumentos.Count + 4), , xlYes)
    tbl.Name = "ResumenColumnas"
    
    ' Ajustar el ancho de las columnas
    For colIndex = 1 To tbl.ListColumns.Count
        If colIndex = 1 Then
            tbl.ListColumns(colIndex).Range.ColumnWidth = 45
        Else
            tbl.ListColumns(colIndex).Range.ColumnWidth = 22
        End If
    Next colIndex
    
    
    Set lastRow1 = tbl.ListRows(tbl.ListRows.Count)
    nombre = Trim(lastRow1.Range(tbl.ListColumns("NOMBRE").Index).Value)
    
    ' Verificar si la longitud de NOMBRE es 0 y eliminar la fila si es así
    If Len(nombre) = 0 Then
        lastRow1.Delete
    End If
    
   ' Centrando los encabezados y los datos
    headerRange.HorizontalAlignment = xlCenter
    headerRange.VerticalAlignment = xlCenter
    headerRange.WrapText = True
    headerRange.RowHeight = 45
    
    tbl.Range.HorizontalAlignment = xlCenter
    tbl.Range.VerticalAlignment = xlCenter
    
    ActivaGestionIntegralCampo
    MsgBox "Resumen de habilitaciones completado.", vbInformation
End Sub


Sub ActivaGestionIntegralCampo()
    ' Activar la hoja "GESTION INTEGRAL CAMPO"
    ThisWorkbook.Sheets("GESTION INTEGRAL CAMPO").Activate
End Sub

Sub ActivaPrincipal()
    ' Ocultar la hoja activa
    ActiveSheet.Visible = xlSheetHidden
    
    ' Activar la hoja "PRINCIPAL"
    ThisWorkbook.Sheets("PRINCIPAL").Activate
End Sub


Sub ResumenHabilitacionesPorVehiculo()
    Dim wsDocumentoServicio As Worksheet
    Dim wsVehiculo As Worksheet
    Dim wsVehiculoDocumentos As Worksheet
    Dim wsGestionIntegralVehiculos As Worksheet
    Dim wsDocXServicio As Worksheet 'Modificación
    Dim wsPrincipal As Worksheet
    Dim wsDocPorTipo As Worksheet
    Dim docPorTipoDict As Object
    Dim docPorTipoDictEmp As Object 'Modificación
    Dim dictDocumentos As Object
    Dim dictVehiculos As Object
    Dim celda As Range
    Dim lastRow As Long
    Dim lastRow1 As ListRow
    Dim proyectoCodigo As String
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim vehiculo As Variant
    Dim documento As Variant
    Dim docPorTipoData As Variant
    Dim docXServicioData As Variant
    Dim placa As String
    Dim nombre As String
    Dim tipoVehiculo As String
    Dim encontrado As Boolean
    Dim tbl As ListObject
    Dim headerRange As Range
    Dim keyObligatorio As String
    Dim docObligatorio As String
    Dim fechaActual As Date

    fechaActual = Date

    ' Inicializar diccionarios
    Set dictDocumentos = CreateObject("Scripting.Dictionary")
    Set dictVehiculos = CreateObject("Scripting.Dictionary")
    Set docPorTipoDict = CreateObject("Scripting.Dictionary")
    Set docPorTipoDictEmp = CreateObject("Scripting.Dictionary") 'Modificación

    ' Establecer referencias a las hojas
    Set wsDocumentoServicio = ThisWorkbook.Sheets("REQUISITOS X VEHICULO")
    Set wsVehiculo = ThisWorkbook.Sheets("VEHICULOS")
    Set wsVehiculoDocumentos = ThisWorkbook.Sheets("VEHICULO DOCUMENTOS")
    Set wsGestionIntegralVehiculos = ThisWorkbook.Sheets("GESTION INTEGRAL VEHICULO CAMPO")
    Set wsPrincipal = ThisWorkbook.Sheets("PRINCIPAL")
    Set wsDocXServicio = ThisWorkbook.Sheets("MASTER REQUISITOS VEHICULO") 'Modificación

    'Modificación
    docXServicioData = wsDocXServicio.Range("A2:E" & wsDocXServicio.Cells(wsDocXServicio.Rows.Count, "A").End(xlUp).Row).Value
    For i = LBound(docXServicioData, 1) To UBound(docXServicioData, 1)
        docPorTipoDictEmp(docXServicioData(i, 3)) = docXServicioData(i, 5)
    Next i

    docPorTipoData = wsDocumentoServicio.Range("A2:E" & wsDocumentoServicio.Cells(wsDocumentoServicio.Rows.Count, "A").End(xlUp).Row).Value
    For i = LBound(docPorTipoData, 1) To UBound(docPorTipoData, 1)
        docPorTipoDict(docPorTipoData(i, 5) & "-" & docPorTipoData(i, 1) & "-" & docPorTipoData(i, 2) & "-" & docPorTipoData(i, 3)) = docPorTipoData(i, 4)
    Next i

    ' Obtener el código del proyecto de la celda A2 en la hoja PRINCIPAL
    proyectoCodigo = wsPrincipal.Range("A2").Value

    ' Obtener documentos únicos de la tabla DocumentoPorPuesto
    lastRow = wsDocumentoServicio.Cells(wsDocumentoServicio.Rows.Count, "C").End(xlUp).Row
    For Each celda In wsDocumentoServicio.Range("C2:C" & lastRow)
        If celda.Offset(0, 2).Value = proyectoCodigo And celda.Offset(0, -2).Value <> "EMBARCACIÓN" Then
            documento = celda.Value
            If Not dictDocumentos.Exists(documento) Then
                dictDocumentos.Add documento, documento
            End If
        End If
    Next celda

    ' Obtener vehículos del proyecto de la tabla VehiculoProyecto
    lastRow = wsVehiculo.Cells(wsVehiculo.Rows.Count, "A").End(xlUp).Row
    For Each celda In wsVehiculo.Range("A2:A" & lastRow)
        If celda.Offset(0, 4).Value = proyectoCodigo And celda.Offset(0, 1).Value <> "EMBARCACIÓN" Then
            vehiculo = celda.Value
            If Not dictVehiculos.Exists(vehiculo) Then
                dictVehiculos.Add vehiculo, Array(vehiculo, celda.Offset(0, 1).Value, celda.Offset(0, 2).Value, celda.Offset(0, 5).Value)
            End If
        End If
    Next celda

    ' Limpiar la hoja de destino
    wsGestionIntegralVehiculos.Cells.Clear

    ' Crear encabezados de la tabla en la hoja GESTION INTEGRAL VEHICULOS CAMPO
    wsGestionIntegralVehiculos.Range("A1").Value = "PLACA"
    wsGestionIntegralVehiculos.Range("B1").Value = "SERVICIO"
    wsGestionIntegralVehiculos.Range("C1").Value = "TIPO"
    wsGestionIntegralVehiculos.Range("D1").Value = "ESTADO"
    colIndex = 5

    For Each documento In dictDocumentos.keys
        protocolo = UCase(docPorTipoDictEmp(documento))
        
        ' Escribe el valor del documento en la celda
        wsGestionIntegralVehiculos.Cells(1, colIndex).Value = documento
        
        ' Verifica si el protocolo es distinto de "TEMA"
        If protocolo <> "TEMA" Then
            ' Si el protocolo es diferente de "TEMA", el fondo es naranja
            wsGestionIntegralVehiculos.Cells(1, colIndex).Interior.Color = RGB(0, 180, 0) ' Verde
        Else
            ' Si el protocolo es igual a "TEMA", el fondo es azul
           wsGestionIntegralVehiculos.Cells(1, colIndex).Interior.Color = RGB(21, 96, 130) ' Azul personalizado de la imagen
       End If
        
        ' Incrementa el índice de la columna
        colIndex = colIndex + 1
    Next documento

  
  ' Llenar la tabla con los vehículos y sus documentos
    rowIndex = 2
    For Each vehiculo In dictVehiculos.keys
        wsGestionIntegralVehiculos.Cells(rowIndex, 1).Value = dictVehiculos(vehiculo)(0)
        wsGestionIntegralVehiculos.Cells(rowIndex, 2).Value = dictVehiculos(vehiculo)(1)
        wsGestionIntegralVehiculos.Cells(rowIndex, 3).Value = dictVehiculos(vehiculo)(2)
        wsGestionIntegralVehiculos.Cells(rowIndex, 4).Value = dictVehiculos(vehiculo)(3)
        servicioVehiculo = dictVehiculos(vehiculo)(1)
        tipoVehiculo = dictVehiculos(vehiculo)(2)
        placa = dictVehiculos(vehiculo)(0)
        
        ' Verificar documentos presentados
        colIndex = 5
        For Each documento In dictDocumentos.keys
            encontrado = False
            encontrado1 = False
            observacion = ""
            lastRow = wsVehiculoDocumentos.Cells(wsVehiculoDocumentos.Rows.Count, "A").End(xlUp).Row
            For Each celda In wsVehiculoDocumentos.Range("A2:A" & lastRow)
                keyObligatorio = proyectoCodigo & "-" & servicioVehiculo & "-" & tipoVehiculo & "-" & documento
                docObligatorio = docPorTipoDict(keyObligatorio)
                If celda.Offset(0, 1).Value = documento And celda.Offset(0, 0).Value = placa Then
                    observacion = celda.Offset(0, 4).Value 'Modificación
                End If
                If docObligatorio = "SI" Then
                    EstadoDoc = ""
                    If celda.Value = placa And celda.Offset(0, 1).Value = documento Then
                        If celda.Offset(0, 3).Value < Date Then
                            EstadoDoc = "VENCIDO"
                        End If
                        encontrado = True
                        Exit For
                    End If
                Else
                    docObligatorio = "NO"
                    EstadoDoc = ""
                    If celda.Value = placa And celda.Offset(0, 1).Value = documento Then
                        If celda.Offset(0, 3).Value < Date Then
                            EstadoDoc = "VENCIDO"
                        Else
                            EstadoDoc = "VÁLIDO"
                        End If
                        encontrado = True
                        Exit For
                    End If
                End If
            Next celda
            
            If encontrado And docObligatorio = "SI" Then
                 If EstadoDoc = "" And observacion <> "APROBADO POR GERENCIA" And observacion <> "NO RESTRICTIVO" And observacion <> "NO APLICA" And observacion <> "OBSERVADO" Then
                    wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "X"
                 Else
                    If observacion = "APROBADO POR GERENCIA" Then
                        wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "AGA"
                    Else
                        If observacion = "NO RESTRICTIVO" Then
                            wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "NR"
                        Else
                            If observacion = "NO APLICA" Then
                                wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "NA"
                            Else
                                If observacion = "OBSERVADO" Then
                                    wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "OBSERVADO"
                                Else
                                    wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "VENCIDO"
                                End If
                            End If
                        End If
                    End If
                 End If
            Else
                If docObligatorio = "NO" Then
                    If EstadoDoc = "" Or EstadoDoc = "VENCIDO" Then
                        wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "NA"
                    Else
                        If observacion = "APROBADO POR GERENCIA" Then
                            wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "AGA"
                        Else
                            If observacion = "NO RESTRICTIVO" Then
                                wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "NR"
                            Else
                                If observacion = "NO APLICA" Then
                                    wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "NA"
                                Else
                                    If observacion = "OBSERVADO" Then
                                        wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "OBSERVADO"
                                    Else
                                        wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "X"
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    If (encontrado = False Or encontrado = Falso) And observacion <> "APROBADO POR GERENCIA" And observacion <> "NO RESTRICTIVO" And observacion <> "NO APLICA" And observacion <> "OBSERVADO" Then
                        wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "FALTA"
                    Else
                    'Modificación
                        If observacion = "APROBADO POR GERENCIA" Then
                            wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "AGA"
                        Else
                            If observacion = "NO RESTRICTIVO" Then
                                wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "NR"
                            Else
                                If observacion = "NO APLICA" Then
                                    wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "NA"
                                Else
                                    If observacion = "OBSERVADO" Then
                                        wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "OBSERVADO"
                                    Else
                                        wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "FALTA"
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            ' Aplica formato condicional si el valor de la celda es "FALTA"
            If wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "FALTA" Then
                With wsGestionIntegralVehiculos.Cells(rowIndex, colIndex)
                    .Interior.Color = RGB(255, 0, 0) ' Fondo rojo
                    .Font.Color = RGB(255, 255, 255) ' Letras blancas
                End With
            ElseIf wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "VENCIDO" Then
                With wsGestionIntegralVehiculos.Cells(rowIndex, colIndex)
                    .Interior.Color = RGB(255, 0, 0) ' Fondo rojo
                    .Font.Color = RGB(0, 0, 0) ' Letras blancas
                End With
            ElseIf wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "OBSERVADO" Then
                With wsGestionIntegralVehiculos.Cells(rowIndex, colIndex)
                    .Interior.Color = RGB(128, 0, 128) ' Fondo rojo
                    .Font.Color = RGB(255, 255, 255) ' Letras blancas
                End With
            End If

            colIndex = colIndex + 1
        Next documento
        
        rowIndex = rowIndex + 1
    Next vehiculo

    ' Eliminar la última fila si está vacía
    If Application.WorksheetFunction.CountA(wsGestionIntegralVehiculos.Rows(rowIndex - 1)) = 0 Then
        wsGestionIntegralVehiculos.Rows(rowIndex - 1).Delete
        rowIndex = rowIndex - 1
    End If

    ' Crear tabla
    Set headerRange = wsGestionIntegralVehiculos.Range(wsGestionIntegralVehiculos.Cells(1, 1), wsGestionIntegralVehiculos.Cells(1, dictDocumentos.Count + 4))
    Set tbl = wsGestionIntegralVehiculos.ListObjects.Add(xlSrcRange, headerRange.Resize(rowIndex, dictDocumentos.Count + 4), , xlYes)
    tbl.Name = "ResumenColumnasVehiculos"
    
    ' Ajustar el ancho de las columnas
    For colIndex = 1 To tbl.ListColumns.Count
        If colIndex = 2 Then
            tbl.ListColumns(colIndex).Range.ColumnWidth = 45
        Else
            tbl.ListColumns(colIndex).Range.ColumnWidth = 22
        End If
    Next colIndex
    
    Set lastRow1 = tbl.ListRows(tbl.ListRows.Count)
    placa = Trim(lastRow1.Range(tbl.ListColumns("PLACA").Index).Value)
    
    ' Verificar si la longitud de placa es 0 y eliminar la fila si es así
    If Len(placa) = 0 Then
        lastRow1.Delete
    End If
    
    ' Centrando los encabezados y los datos
    headerRange.HorizontalAlignment = xlCenter
    headerRange.VerticalAlignment = xlCenter
    headerRange.WrapText = True
    headerRange.RowHeight = 45
    
    tbl.Range.HorizontalAlignment = xlCenter
    tbl.Range.VerticalAlignment = xlCenter
    ActivaGestionIntegralVehiculosCampo
    MsgBox "Resumen de habilitaciones de vehículos completado.", vbInformation
End Sub

Sub ActivaGestionIntegralVehiculosCampo()
    ' Activar la hoja "GESTION INTEGRAL VEHICULOS CAMPO"
    ThisWorkbook.Sheets("GESTION INTEGRAL VEHICULO CAMPO").Activate
End Sub

Sub ActivaGestionIntegralEmbarcacionesCampo()
    ' Activar la hoja "GESTION INTEGRAL EMBARCACIONES CAMPO"
    ThisWorkbook.Sheets("GESTION INTEGRAL EMBARCACIONES").Activate
End Sub


Sub ResumenHabilitacionesPorEmbarcacion()
    Dim wsDocumentoServicio As Worksheet
    Dim wsVehiculo As Worksheet
    Dim wsVehiculoDocumentos As Worksheet
    Dim wsGestionIntegralVehiculos As Worksheet
    Dim wsDocXServicio As Worksheet 'Modificación
    Dim wsPrincipal As Worksheet
    Dim wsDocPorTipo As Worksheet
    Dim docPorTipoDict As Object
    Dim docPorTipoDictEmp As Object 'Modificación
    Dim dictDocumentos As Object
    Dim dictVehiculos As Object
    Dim celda As Range
    Dim lastRow As Long
    Dim lastRow1 As ListRow
    Dim proyectoCodigo As String
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim vehiculo As Variant
    Dim documento As Variant
    Dim docPorTipoData As Variant
    Dim docXServicioData As Variant
    Dim placa As String
    Dim nombre As String
    Dim tipoVehiculo As String
    Dim encontrado As Boolean
    Dim tbl As ListObject
    Dim headerRange As Range
    Dim keyObligatorio As String
    Dim docObligatorio As String
    Dim fechaActual As Date

    fechaActual = Date

    ' Inicializar diccionarios
    Set dictDocumentos = CreateObject("Scripting.Dictionary")
    Set dictVehiculos = CreateObject("Scripting.Dictionary")
    Set docPorTipoDict = CreateObject("Scripting.Dictionary")
    Set docPorTipoDictEmp = CreateObject("Scripting.Dictionary") 'Modificación

    ' Establecer referencias a las hojas
    Set wsDocumentoServicio = ThisWorkbook.Sheets("REQUISITOS X VEHICULO")
    Set wsVehiculo = ThisWorkbook.Sheets("VEHICULOS")
    Set wsVehiculoDocumentos = ThisWorkbook.Sheets("VEHICULO DOCUMENTOS")
    Set wsGestionIntegralVehiculos = ThisWorkbook.Sheets("GESTION INTEGRAL EMBARCACIONES")
    Set wsPrincipal = ThisWorkbook.Sheets("PRINCIPAL")
    Set wsDocXServicio = ThisWorkbook.Sheets("MASTER REQUISITOS VEHICULO") 'Modificación

    'Modificación
    docXServicioData = wsDocXServicio.Range("A2:E" & wsDocXServicio.Cells(wsDocXServicio.Rows.Count, "A").End(xlUp).Row).Value
    For i = LBound(docXServicioData, 1) To UBound(docXServicioData, 1)
        docPorTipoDictEmp(docXServicioData(i, 3)) = docXServicioData(i, 5)
    Next i

    docPorTipoData = wsDocumentoServicio.Range("A2:E" & wsDocumentoServicio.Cells(wsDocumentoServicio.Rows.Count, "A").End(xlUp).Row).Value
    For i = LBound(docPorTipoData, 1) To UBound(docPorTipoData, 1)
        docPorTipoDict(docPorTipoData(i, 5) & "-" & docPorTipoData(i, 1) & "-" & docPorTipoData(i, 2) & "-" & docPorTipoData(i, 3)) = docPorTipoData(i, 4)
    Next i

    ' Obtener el código del proyecto de la celda A2 en la hoja PRINCIPAL
    proyectoCodigo = wsPrincipal.Range("A2").Value

    ' Obtener documentos únicos de la tabla DocumentoPorPuesto
    lastRow = wsDocumentoServicio.Cells(wsDocumentoServicio.Rows.Count, "C").End(xlUp).Row
    For Each celda In wsDocumentoServicio.Range("C2:C" & lastRow)
        If celda.Offset(0, 2).Value = proyectoCodigo And celda.Offset(0, -2).Value = "EMBARCACIÓN" Then
            documento = celda.Value
            If Not dictDocumentos.Exists(documento) Then
                dictDocumentos.Add documento, documento
            End If
        End If
    Next celda

    ' Obtener vehículos del proyecto de la tabla VehiculoProyecto
    lastRow = wsVehiculo.Cells(wsVehiculo.Rows.Count, "A").End(xlUp).Row
    For Each celda In wsVehiculo.Range("A2:A" & lastRow)
        If celda.Offset(0, 4).Value = proyectoCodigo And celda.Offset(0, 1).Value = "EMBARCACIÓN" Then
            vehiculo = celda.Value
            If Not dictVehiculos.Exists(vehiculo) Then
                dictVehiculos.Add vehiculo, Array(vehiculo, celda.Offset(0, 1).Value, celda.Offset(0, 2).Value, celda.Offset(0, 5).Value)
            End If
        End If
    Next celda

    ' Limpiar la hoja de destino
    wsGestionIntegralVehiculos.Cells.Clear

    ' Crear encabezados de la tabla en la hoja GESTION INTEGRAL VEHICULOS CAMPO
    wsGestionIntegralVehiculos.Range("A1").Value = "PLACA"
    wsGestionIntegralVehiculos.Range("B1").Value = "SERVICIO"
    wsGestionIntegralVehiculos.Range("C1").Value = "TIPO"
    wsGestionIntegralVehiculos.Range("D1").Value = "ESTADO"
    colIndex = 5

    For Each documento In dictDocumentos.keys
        protocolo = UCase(docPorTipoDictEmp(documento))
        
        ' Escribe el valor del documento en la celda
        wsGestionIntegralVehiculos.Cells(1, colIndex).Value = documento
        
        ' Verifica si el protocolo es distinto de "TEMA"
        If protocolo <> "TEMA" Then
            ' Si el protocolo es diferente de "TEMA", el fondo es naranja
            wsGestionIntegralVehiculos.Cells(1, colIndex).Interior.Color = RGB(0, 180, 0) ' Verde
        Else
            ' Si el protocolo es igual a "TEMA", el fondo es azul
           wsGestionIntegralVehiculos.Cells(1, colIndex).Interior.Color = RGB(21, 96, 130) ' Azul personalizado de la imagen
       End If
        
        ' Incrementa el índice de la columna
        colIndex = colIndex + 1
    Next documento

  
  ' Llenar la tabla con los vehículos y sus documentos
    rowIndex = 2
    For Each vehiculo In dictVehiculos.keys
        wsGestionIntegralVehiculos.Cells(rowIndex, 1).Value = dictVehiculos(vehiculo)(0)
        wsGestionIntegralVehiculos.Cells(rowIndex, 2).Value = dictVehiculos(vehiculo)(1)
        wsGestionIntegralVehiculos.Cells(rowIndex, 3).Value = dictVehiculos(vehiculo)(2)
        wsGestionIntegralVehiculos.Cells(rowIndex, 4).Value = dictVehiculos(vehiculo)(3)
        servicioVehiculo = dictVehiculos(vehiculo)(1)
        tipoVehiculo = dictVehiculos(vehiculo)(2)
        placa = dictVehiculos(vehiculo)(0)
        
        ' Verificar documentos presentados
        colIndex = 5
        For Each documento In dictDocumentos.keys
            encontrado = False
            encontrado1 = False
            observacion = ""
            lastRow = wsVehiculoDocumentos.Cells(wsVehiculoDocumentos.Rows.Count, "A").End(xlUp).Row
            For Each celda In wsVehiculoDocumentos.Range("A2:A" & lastRow)
                keyObligatorio = proyectoCodigo & "-" & servicioVehiculo & "-" & tipoVehiculo & "-" & documento
                docObligatorio = docPorTipoDict(keyObligatorio)
                If celda.Offset(0, 1).Value = documento And celda.Offset(0, 0).Value = placa Then
                    observacion = celda.Offset(0, 4).Value 'Modificación
                End If
                If docObligatorio = "SI" Then
                    EstadoDoc = ""
                    If celda.Value = placa And celda.Offset(0, 1).Value = documento Then
                        If celda.Offset(0, 3).Value < Date Then
                            EstadoDoc = "VENCIDO"
                        End If
                        encontrado = True
                        Exit For
                    End If
                Else
                    docObligatorio = "NO"
                    EstadoDoc = ""
                    If celda.Value = placa And celda.Offset(0, 1).Value = documento Then
                        If celda.Offset(0, 3).Value < Date Then
                            EstadoDoc = "VENCIDO"
                        Else
                            EstadoDoc = "VÁLIDO"
                        End If
                        encontrado = True
                        Exit For
                    End If
                End If
            Next celda
            
            If encontrado And docObligatorio = "SI" Then
                 If EstadoDoc = "" And observacion <> "APROBADO POR GERENCIA" And observacion <> "NO RESTRICTIVO" And observacion <> "NO APLICA" And observacion <> "OBSERVADO" Then
                    wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "X"
                 Else
                    If observacion = "APROBADO POR GERENCIA" Then
                        wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "AGA"
                    Else
                        If observacion = "NO RESTRICTIVO" Then
                            wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "NR"
                        Else
                            If observacion = "NO APLICA" Then
                                wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "NA"
                            Else
                                If observacion = "OBSERVADO" Then
                                    wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "OBSERVADO"
                                Else
                                    wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "VENCIDO"
                                End If
                            End If
                        End If
                    End If
                 End If
            Else
                If docObligatorio = "NO" Then
                    If EstadoDoc = "" Or EstadoDoc = "VENCIDO" Then
                        wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "NA"
                    Else
                        If observacion = "APROBADO POR GERENCIA" Then
                            wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "AGA"
                        Else
                            If observacion = "NO RESTRICTIVO" Then
                                wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "NR"
                            Else
                                If observacion = "NO APLICA" Then
                                    wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "NA"
                                Else
                                    If observacion = "OBSERVADO" Then
                                        wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "OBSERVADO"
                                    Else
                                        wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "X"
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    If (encontrado = False Or encontrado = Falso) And observacion <> "APROBADO POR GERENCIA" And observacion <> "NO RESTRICTIVO" And observacion <> "NO APLICA" And observacion <> "OBSERVADO" Then
                        wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "FALTA"
                    Else
                    'Modificación
                        If observacion = "APROBADO POR GERENCIA" Then
                            wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "AGA"
                        Else
                            If observacion = "NO RESTRICTIVO" Then
                                wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "NR"
                            Else
                                If observacion = "NO APLICA" Then
                                    wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "NA"
                                Else
                                    If observacion = "OBSERVADO" Then
                                        wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "OBSERVADO"
                                    Else
                                        wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "FALTA"
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            ' Aplica formato condicional si el valor de la celda es "FALTA"
            If wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "FALTA" Then
                With wsGestionIntegralVehiculos.Cells(rowIndex, colIndex)
                    .Interior.Color = RGB(255, 0, 0) ' Fondo rojo
                    .Font.Color = RGB(255, 255, 255) ' Letras blancas
                End With
            ElseIf wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "VENCIDO" Then
                With wsGestionIntegralVehiculos.Cells(rowIndex, colIndex)
                    .Interior.Color = RGB(255, 0, 0) ' Fondo rojo
                    .Font.Color = RGB(0, 0, 0) ' Letras blancas
                End With
            ElseIf wsGestionIntegralVehiculos.Cells(rowIndex, colIndex).Value = "OBSERVADO" Then
                With wsGestionIntegralVehiculos.Cells(rowIndex, colIndex)
                    .Interior.Color = RGB(128, 0, 128) ' Fondo rojo
                    .Font.Color = RGB(255, 255, 255) ' Letras blancas
                End With
            End If

            colIndex = colIndex + 1
        Next documento
        
        rowIndex = rowIndex + 1
    Next vehiculo

    ' Eliminar la última fila si está vacía
    If Application.WorksheetFunction.CountA(wsGestionIntegralVehiculos.Rows(rowIndex - 1)) = 0 Then
        wsGestionIntegralVehiculos.Rows(rowIndex - 1).Delete
        rowIndex = rowIndex - 1
    End If

    ' Crear tabla
    Set headerRange = wsGestionIntegralVehiculos.Range(wsGestionIntegralVehiculos.Cells(1, 1), wsGestionIntegralVehiculos.Cells(1, dictDocumentos.Count + 4))
    Set tbl = wsGestionIntegralVehiculos.ListObjects.Add(xlSrcRange, headerRange.Resize(rowIndex, dictDocumentos.Count + 4), , xlYes)
    tbl.Name = "ResumenColumnasVehiculos"
    
    ' Ajustar el ancho de las columnas
    For colIndex = 1 To tbl.ListColumns.Count
        If colIndex = 2 Then
            tbl.ListColumns(colIndex).Range.ColumnWidth = 45
        Else
            tbl.ListColumns(colIndex).Range.ColumnWidth = 22
        End If
    Next colIndex
    
    Set lastRow1 = tbl.ListRows(tbl.ListRows.Count)
    placa = Trim(lastRow1.Range(tbl.ListColumns("PLACA").Index).Value)
    
    ' Verificar si la longitud de placa es 0 y eliminar la fila si es así
    If Len(placa) = 0 Then
        lastRow1.Delete
    End If
    
    ' Centrando los encabezados y los datos
    headerRange.HorizontalAlignment = xlCenter
    headerRange.VerticalAlignment = xlCenter
    headerRange.WrapText = True
    headerRange.RowHeight = 45
    
    tbl.Range.HorizontalAlignment = xlCenter
    tbl.Range.VerticalAlignment = xlCenter
    ActivaGestionIntegralEmbarcacionesCampo
    MsgBox "Resumen de habilitaciones de embarcaciones completado.", vbInformation
End Sub

Sub BackupArchivosEmbarcaciones()
    Dim ws As Worksheet
    Dim wsPrincipal As Worksheet
    Dim tbl As ListObject
    Dim fd As FileDialog
    Dim destFolder As String
    Dim projectFolder As String
    Dim vehiculoFolder As String
    Dim companyFolder As String
    Dim docFolder As String
    Dim projectCode As String
    Dim companyName As String
    Dim docID As String
    Dim filePath As String
    Dim fullFilePath As String
    Dim carpetaBase As String
    Dim lastRow As Long
    Dim i As Long
    Dim fso As Object
    Dim currentPath As String
    Dim destino As String

    ' Crear objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Obtener la ruta del directorio actual
    currentPath = fso.GetAbsolutePathName(".")
    Debug.Print "Directorio actual: " & currentPath

    ' Configurar la hoja de trabajo y la tabla
    Set ws = ThisWorkbook.Sheets("DASHBOARD EMBARCACIONES")
    Set wsPrincipal = ThisWorkbook.Sheets("PRINCIPAL")
    Set tbl = ws.ListObjects("DashboardEmbarcacion")
    
    ' Crear ventana de diálogo para seleccionar la carpeta de destino
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Seleccionar carpeta de destino para el backup"
        If .Show = -1 Then
            destFolder = .SelectedItems(1)
        Else
            MsgBox "No se seleccionó ninguna carpeta. Proceso cancelado.", vbExclamation
            Exit Sub
        End If
    End With
    
    Debug.Print "Carpeta de destino seleccionada: " & destFolder

    carpetaBase = wsPrincipal.Cells(6, 2).Value
    Debug.Print "Carpeta base: " & carpetaBase
    
    ' Recorrer cada fila de la tabla
    lastRow = tbl.ListRows.Count
    For i = 1 To lastRow
        ' Leer valores de las columnas
        projectCode = tbl.ListColumns("CÓDIGO PROYECTO").DataBodyRange(i, 1).Value
        companyName = tbl.ListColumns("EMPRESA").DataBodyRange(i, 1).Value
        docID = tbl.ListColumns("PLACA").DataBodyRange(i, 1).Value
        filePath = tbl.ListColumns("LINK").DataBodyRange(i, 1).Hyperlinks(1).Address
        
        Debug.Print "Procesando fila: " & i
        Debug.Print "Código de proyecto: " & projectCode
        Debug.Print "Nombre de la empresa: " & companyName
        Debug.Print "Placa Vehículo: " & docID
        Debug.Print "Ruta del archivo: " & filePath
        
        ' Obtener la ruta completa si es relativa
        If Left(filePath, 2) = ".." Or Left(filePath, 1) = "\" Then
            fullFilePath = fso.GetAbsolutePathName(fso.BuildPath(currentPath, filePath))
        Else
            fullFilePath = filePath
        End If
        
        Debug.Print "Ruta completa del archivo: " & fullFilePath

        ' Crear rutas de las carpetas
        projectFolder = destFolder & "\" & projectCode
        vehiculoFolder = projectFolder & "\EMBARCACIONES"
        companyFolder = vehiculoFolder & "\" & companyName
        docFolder = companyFolder & "\" & docID
        
        ' Crear las carpetas si no existen
        CrearCarpetas projectFolder
        CrearCarpetas vehiculoFolder
        CrearCarpetas companyFolder
        CrearCarpetas docFolder
        
        destino = docFolder & "\" & fso.GetFileName(fullFilePath)

        CrearCarpetas destino
        ' Verificar la ruta de destino
        Debug.Print "Ruta de destino del archivo: " & destino
        
        Set folder = fso.GetFolder(fullFilePath)
        If folder.Files.Count > 0 Then
            ' Copiar todos los archivos de la carpeta de origen a la carpeta de destino
                For Each file In folder.Files
                    Debug.Print "Copiando archivo desde: " & file.Path & " a " & destino & "\" & file.Name
                    fso.CopyFile Source:=file.Path, Destination:=destino & "\" & file.Name
                    Debug.Print "Archivo copiado: " & destino & "\" & file.Name
                Next file
        Else
            Debug.Print "Archivo no encontrado: " & fullFilePath
        End If
    Next i
    
    MsgBox "Backup completado exitosamente.", vbInformation
End Sub

Sub EliminarCarpetas()
    EliminarCarpetasInnecesarias ("")
End Sub


Sub EliminarCarpetasInnecesarias(Optional Logical As Variant)
    Dim wsPrincipal As Worksheet, wsPersonal As Worksheet, wsVehiculos As Worksheet
    Dim wsDocumentosPuestos As Worksheet, wsMasterDocumentos As Worksheet
    Dim wsRequisitosVehiculos As Worksheet, wsMasterRequisitos As Worksheet
    Dim carpetaRaiz As String, carpetaRecursos As String
    Dim fso As Object, carpetasValidas As Object, proyectoCodigos As Object
    Dim personalData() As String, vehiculosData() As String
    Dim documentosPuestosData() As String, masterDocumentosData() As String
    Dim requisitosVehiculosData() As String, masterRequisitosData() As String
    Dim i As Long, j As Long, k As Long, proyectoCodigo As Variant
    Dim documentoNombre As String, vehiculoNombre As String, tipoDocumento As String
    Dim carpetaProyecto As String, carpetaHabilitaciones As String, carpetaPersonal As String, carpetaVehiculos As String
    Dim carpetaDocumento As String, carpetaRequisito As String

    ' Optimización: Desactivar actualización de pantalla y cálculos
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Definir ruta del archivo de depuración
    rutaArchivoDepuracion = "C:\Users\lhoyos\Downloads\Depuracion_Carpetas.txt" ' Cambia la ruta si es necesario
    archivo = FreeFile

    ' Crear o sobrescribir el archivo de depuración
    Open rutaArchivoDepuracion For Output As archivo
    Print #archivo, "Inicio de Depuración - Eliminación de Carpetas Innecesarias"
    Print #archivo, "------------------------------------------------------------"

    ' Establecer hojas
    Set wsPrincipal = ThisWorkbook.Sheets("PRINCIPAL")
    Set wsPersonal = ThisWorkbook.Sheets("PERSONAL")
    Set wsVehiculos = ThisWorkbook.Sheets("VEHICULOS")
    Set wsDocumentosPuestos = ThisWorkbook.Sheets("DOCUMENTO X PUESTO")
    Set wsMasterDocumentos = ThisWorkbook.Sheets("MASTER DOCUMENTOS PERSONAL")
    Set wsRequisitosVehiculos = ThisWorkbook.Sheets("REQUISITOS X VEHICULO")
    Set wsMasterRequisitos = ThisWorkbook.Sheets("MASTER REQUISITOS VEHICULO")

    ' Obtener carpeta raíz
    carpetaRaiz = wsPrincipal.Range("B6").Value
    If Right(carpetaRaiz, 1) <> "\" Then carpetaRaiz = carpetaRaiz & "\"
    carpetaRecursos = carpetaRaiz & "RECURSOS"

    ' Crear objetos
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set carpetasValidas = CreateObject("Scripting.Dictionary")
    Set proyectoCodigos = CreateObject("Scripting.Dictionary")

    ' Cargar datos en Arrays optimizados en lugar de `Variant`
    personalData = RangoAArray(wsPersonal.Range("A2:G" & wsPersonal.Cells(wsPersonal.Rows.Count, 1).End(xlUp).Row))
    vehiculosData = RangoAArray(wsVehiculos.Range("A2:F" & wsVehiculos.Cells(wsVehiculos.Rows.Count, 1).End(xlUp).Row))
    documentosPuestosData = RangoAArray(wsDocumentosPuestos.Range("A1:D" & wsDocumentosPuestos.Cells(wsDocumentosPuestos.Rows.Count, 1).End(xlUp).Row))
    masterDocumentosData = RangoAArray(wsMasterDocumentos.Range("A1:F" & wsMasterDocumentos.Cells(wsMasterDocumentos.Rows.Count, 1).End(xlUp).Row))
    requisitosVehiculosData = RangoAArray(wsRequisitosVehiculos.Range("A1:E" & wsRequisitosVehiculos.Cells(wsRequisitosVehiculos.Rows.Count, 1).End(xlUp).Row))
    masterRequisitosData = RangoAArray(wsMasterRequisitos.Range("A1:G" & wsMasterRequisitos.Cells(wsMasterRequisitos.Rows.Count, 1).End(xlUp).Row))
    
    ' Obtener códigos de proyectos únicos
    On Error Resume Next
    For i = 1 To UBound(personalData, 1)
        proyectoCodigos.Add personalData(i, 5), CStr(personalData(i, 5))
        Print #archivo, "- " & personalData(i, 5)
        Debug.Print "- " & personalData(i, 5)
    Next i
    For i = 1 To UBound(vehiculosData, 1)
        proyectoCodigos.Add vehiculosData(i, 5), CStr(vehiculosData(i, 5))
        Print #archivo, "- " & vehiculosData(i, 5)
        Debug.Print "- " & vehiculosData(i, 5)
    Next i
    On Error GoTo 0

    ' Generar lista de carpetas válidas
    For Each proyectoCodigo In proyectoCodigos.keys
        carpetaProyecto = carpetaRaiz & proyectoCodigo
        carpetaHabilitaciones = carpetaProyecto & "\HABILITACIONES"

        ' Carpetas de Personal
        For i = 1 To UBound(personalData, 1)
            If CStr(personalData(i, 5)) = proyectoCodigo Then
                documentoNombre = Trim(personalData(i, 1))
                carpetaPersonal = carpetaHabilitaciones & "\PERSONAL\" & documentoNombre
                carpetasValidas(carpetaPersonal) = True
                carpetasValidas(carpetaRecursos & "\PERSONAL\" & documentoNombre) = True

                Print #archivo, "Carpeta Personal Agregada: " & carpetaPersonal
                Debug.Print "Carpeta Personal Agregada: " & carpetaPersonal
                Print #archivo, "Carpeta Personal Agregada: " & carpetaRecursos & "\PERSONAL\" & documentoNombre
                Debug.Print "Carpeta Personal Agregada: " & carpetaRecursos & "\PERSONAL\" & documentoNombre

                ' Buscar documentos relacionados al puesto
                For j = 2 To UBound(documentosPuestosData, 1)
                    If documentosPuestosData(j, 1) = personalData(i, 3) And documentosPuestosData(j, 4) = proyectoCodigo Then
                        tipoDocumento = documentosPuestosData(j, 2)
                        For k = 2 To UBound(masterDocumentosData, 1)
                            If masterDocumentosData(k, 1) = personalData(i, 3) And _
                               masterDocumentosData(k, 2) = tipoDocumento Then

                                carpetaDocumento = carpetaPersonal & "\" & masterDocumentosData(k, 3)
                                If Not carpetasValidas.Exists(carpetaDocumento) Then
                                    carpetasValidas(carpetaDocumento) = True
                                    Print #archivo, "Documento Personal Relacionado: " & carpetaDocumento
                                    Debug.Print "Documento Personal Relacionado: " & carpetaDocumento
                                End If
                            
                                carpetaDocumento = carpetaRecursos & "\PERSONAL\" & documentoNombre
                                If Not carpetasValidas.Exists(carpetaDocumento) Then
                                    carpetasValidas(carpetaDocumento) = True
                                    Print #archivo, "Documento Personal Relacionado: " & carpetaDocumento
                                    Debug.Print "Documento Personal Relacionado: " & carpetaDocumento
                                End If
                            
                                carpetaDocumento = carpetaRecursos & "\PERSONAL\" & documentoNombre & "\" & masterDocumentosData(k, 3)
                                If Not carpetasValidas.Exists(carpetaDocumento) Then
                                    carpetasValidas(carpetaDocumento) = True
                                    Print #archivo, "Documento Personal Relacionado: " & carpetaDocumento
                                    Debug.Print "Documento Personal Relacionado: " & carpetaDocumento
                                End If
                            End If
                        Next k
                    End If
                Next j
            End If
        Next i

        ' Carpetas de Vehículos
        For i = 1 To UBound(vehiculosData, 1)
            If CStr(vehiculosData(i, 5)) = proyectoCodigo Then
                vehiculoNombre = vehiculosData(i, 1)
                carpetaVehiculos = carpetaHabilitaciones & "\VEHICULOS\" & vehiculoNombre
                carpetasValidas(carpetaVehiculos) = True
                
                Print #archivo, "Carpeta Vehículo Agregada: " & carpetaVehiculos
                Debug.Print "Carpeta Vehículo Agregada: " & carpetaVehiculos
                
                carpetasValidas(carpetaRecursos & "\VEHICULOS\" & vehiculoNombre) = True

                Print #archivo, "Carpeta Vehículo Agregada: " & carpetaRecursos & "\VEHICULOS\" & vehiculoNombre
                Debug.Print "Carpeta Vehículo Agregada: " & carpetaRecursos & "\VEHICULOS\" & vehiculoNombre

                ' Buscar requisitos relacionados al vehículo
                For j = 2 To UBound(requisitosVehiculosData, 1)
                    If requisitosVehiculosData(j, 1) = vehiculosData(i, 2) And _
                       requisitosVehiculosData(j, 2) = vehiculosData(i, 3) And _
                       requisitosVehiculosData(j, 5) = proyectoCodigo Then
                        For k = 2 To UBound(masterRequisitosData, 1)
                            If masterRequisitosData(k, 1) = vehiculosData(i, 2) And _
                               masterRequisitosData(k, 2) = vehiculosData(i, 3) And _
                               masterRequisitosData(k, 3) = requisitosVehiculosData(j, 3) Then

                                carpetaRequisito = carpetaVehiculos & "\" & masterRequisitosData(k, 4)
                                If Not carpetasValidas.Exists(carpetaRequisito) Then
                                    carpetasValidas(carpetaRequisito) = True
                                    Print #archivo, "Requisito Vehículo Relacionado: " & carpetaRequisito
                                    Debug.Print "Requisito Vehículo Relacionado: " & carpetaRequisito
                                End If
                            
                                carpetaRequisito = carpetaRecursos & "\VEHICULOS\" & vehiculoNombre
                                If Not carpetasValidas.Exists(carpetaRequisito) Then
                                    carpetasValidas(carpetaRequisito) = True
                                    Print #archivo, "Requisito Vehículo Relacionado: " & carpetaRequisito
                                    Debug.Print "Requisito Vehículo Relacionado: " & carpetaRequisito
                                End If
                            
                                carpetaRequisito = carpetaRecursos & "\VEHICULOS\" & vehiculoNombre & "\" & masterRequisitosData(k, 4)
                                If Not carpetasValidas.Exists(carpetaRequisito) Then
                                    carpetasValidas(carpetaRequisito) = True
                                    Print #archivo, "Requisito Vehículo Relacionado: " & carpetaRequisito
                                    Debug.Print "Requisito Vehículo Relacionado: " & carpetaRequisito
                                End If
                            End If
                        Next k
                    End If
                Next j
            End If
        Next i
    Next proyectoCodigo

    ' Eliminar carpetas innecesarias en HABILITACIONES
    For Each proyectoCodigo In proyectoCodigos.keys
        carpetaProyecto = carpetaRaiz & proyectoCodigo
        carpetaHabilitaciones = carpetaProyecto & "\HABILITACIONES"
        EliminarCarpetasNoValidas fso, carpetasValidas, carpetaHabilitaciones & "\PERSONAL"
        EliminarCarpetasNoValidas fso, carpetasValidas, carpetaHabilitaciones & "\VEHICULOS"
    Next proyectoCodigo
    
    ' Eliminar carpetas innecesarias en RECURSOS
    EliminarCarpetasNoValidas fso, carpetasValidas, carpetaRecursos & "\PERSONAL"
    EliminarCarpetasNoValidas fso, carpetasValidas, carpetaRecursos & "\VEHICULOS"

    ' Cerrar archivo de depuración
    Print #archivo, "------------------------------------------------------------"
    Print #archivo, "Fin de Depuración"
    Close archivo
    
    ' Restaurar configuración de Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    ' Mensaje de confirmación
    If IsMissing(Logical) Or IsEmpty(Logical) Or Logical = "" Then
        MsgBox "Carpetas innecesarias eliminadas correctamente."
    End If
End Sub
' ?? Función para convertir un Rango a Array (más rápido que `Variant`)
Function RangoAArray(rng As Range) As String()
    Dim data As Variant, arr() As String
    Dim i As Long, j As Long
    data = rng.Value ' Carga el rango en un `Variant`
    
    ' Redimensionar el Array con el mismo tamaño del `Variant`
    ReDim arr(LBound(data, 1) To UBound(data, 1), LBound(data, 2) To UBound(data, 2))
    
    ' Copiar datos del `Variant` al `Array`
    For i = LBound(data, 1) To UBound(data, 1)
        For j = LBound(data, 2) To UBound(data, 2)
            arr(i, j) = CStr(data(i, j)) ' Convertir a String para evitar errores
        Next j
    Next i
    
    RangoAArray = arr
End Function
' ?? Optimización de eliminación de carpetas
Sub EliminarCarpetasNoValidas(fso As Object, carpetasValidas As Object, rutaBase As String)
    Dim carpeta As Object, subcarpeta As Object
    If Not fso.FolderExists(rutaBase) Then Exit Sub

    For Each carpeta In fso.GetFolder(rutaBase).SubFolders
        If Not carpetasValidas.Exists(carpeta.Path) Then
            On Error Resume Next
            SetAttr carpeta.Path, vbNormal ' Desbloquea la carpeta
            fso.DeleteFolder carpeta.Path, True
            On Error GoTo 0
        Else
            For Each subcarpeta In carpeta.SubFolders
                'If carpeta.Path = "C:\Users\lhoyos\Tema Litoclean\GESPRO - Documentos\General\10706 HABILITACIONES\General\RECURSOS\PERSONAL\41031557" Then Stop
                'If carpeta.Path = "C:\Users\lhoyos\Tema Litoclean\GESPRO - Documentos\General\10706 HABILITACIONES\General\RECURSOS\PERSONAL\41031557\ANPE-R" Then Stop
                If Not carpetasValidas.Exists(subcarpeta.Path) Then
                    On Error Resume Next
                    SetAttr subcarpeta.Path, vbNormal
                    fso.DeleteFolder subcarpeta.Path, True
                    On Error GoTo 0
                End If
            Next subcarpeta
        End If
    Next carpeta
End Sub





Sub EliminarCarpetasInnecesariasSUPERADO(Optional Logical As Variant)
    Dim wsPrincipal As Worksheet, wsPersonal As Worksheet, wsDocumentosPuestos As Worksheet
    Dim wsVehiculos As Worksheet, wsRequisitosVehiculos As Worksheet
    Dim wsMasterDocumentos As Worksheet, wsMasterRequisitos As Worksheet
    Dim wsRutasValidas As Worksheet
    Dim proyectoCodigos As Collection, proyectoCodigo As Variant
    Dim carpetaRaiz As String, carpetaProyecto As String
    Dim carpetaHabilitaciones As String, carpetaRecursos As String
    Dim carpetaPersonal As String, carpetaVehiculos As String
    Dim personalData As Variant, vehiculosData As Variant, documentosPuestosData As Variant
    Dim masterDocumentosData As Variant, requisitosVehiculosData As Variant, masterRequisitosData As Variant
    Dim i As Long, j As Long, k As Long
    Dim documentoNombre As String, vehiculoNombre As String
    Dim carpetaActual As String
    Dim fso As Object
    Dim carpeta As Object
    Dim subcarpeta As Object
    Dim carpetasValidas As Collection
    Dim ultimaFila As Long

    ' Establecer las hojas
    Set wsPrincipal = ThisWorkbook.Sheets("PRINCIPAL")
    Set wsPersonal = ThisWorkbook.Sheets("PERSONAL")
    Set wsDocumentosPuestos = ThisWorkbook.Sheets("DOCUMENTO X PUESTO")
    Set wsVehiculos = ThisWorkbook.Sheets("VEHICULOS")
    Set wsRequisitosVehiculos = ThisWorkbook.Sheets("REQUISITOS X VEHICULO")
    Set wsMasterDocumentos = ThisWorkbook.Sheets("MASTER DOCUMENTOS PERSONAL")
    Set wsMasterRequisitos = ThisWorkbook.Sheets("MASTER REQUISITOS VEHICULO")
    'Set wsRutasValidas = ThisWorkbook.Sheets("RutasValidas")

    ' Limpiar la hoja RutasValidas antes de comenzar
    'wsRutasValidas.Cells.Clear

    ' Obtener la carpeta raíz
    carpetaRaiz = wsPrincipal.Range("B6").Value
    
    ' Verificar si la última parte de carpetaRaiz es una barra invertida
    If Right(carpetaRaiz, 1) <> "\" Then
        carpetaRaiz = carpetaRaiz & "\"
    End If
    carpetaRecursos = carpetaRaiz & "RECURSOS"

    ' Cargar los datos en arreglos para acceso más rápido
    personalData = wsPersonal.Range("A2:E" & wsPersonal.Cells(wsPersonal.Rows.Count, "A").End(xlUp).Row).Value
    documentosPuestosData = wsDocumentosPuestos.UsedRange.Value
    masterDocumentosData = wsMasterDocumentos.UsedRange.Value
    vehiculosData = wsVehiculos.Range("A2:E" & wsVehiculos.Cells(wsVehiculos.Rows.Count, "A").End(xlUp).Row).Value
    requisitosVehiculosData = wsRequisitosVehiculos.UsedRange.Value
    masterRequisitosData = wsMasterRequisitos.UsedRange.Value

    ' Crear una colección para almacenar los códigos de proyecto únicos
    Set proyectoCodigos = New Collection

    ' Obtener códigos de proyecto únicos de la hoja PERSONAL
    On Error Resume Next
    For i = 1 To UBound(personalData, 1)
        proyectoCodigos.Add personalData(i, 5), CStr(personalData(i, 5))
    Next i
    On Error GoTo 0

    ' Obtener códigos de proyecto únicos de la hoja VEHICULOS
    On Error Resume Next
    For i = 1 To UBound(vehiculosData, 1)
        proyectoCodigos.Add vehiculosData(i, 5), CStr(vehiculosData(i, 5))
    Next i
    On Error GoTo 0

    ' Crear una colección para almacenar las carpetas válidas
    Set carpetasValidas = New Collection

    ' Loop para cada código de proyecto (Solo para recopilar rutas válidas)
    For Each proyectoCodigo In proyectoCodigos
        carpetaProyecto = carpetaRaiz & proyectoCodigo
        carpetaHabilitaciones = carpetaProyecto & "\HABILITACIONES"
        
        ' Obtener carpetas válidas para Personal
        For i = 1 To UBound(personalData, 1)
            If CStr(personalData(i, 5)) = proyectoCodigo Then
                documentoNombre = Trim(personalData(i, 1))
                carpetaPersonal = carpetaHabilitaciones & "\PERSONAL\" & documentoNombre
                carpetasValidas.Add carpetaPersonal
                ' Añadir la carpeta a la hoja RutasValidas
                'ultimaFila = wsRutasValidas.Cells(wsRutasValidas.Rows.Count, 1).End(xlUp).Row + 1
                'wsRutasValidas.Cells(ultimaFila, 1).Value = carpetaPersonal
                
                ' Buscar documentos relacionados al puesto
                For j = 2 To UBound(documentosPuestosData, 1)
                    If documentosPuestosData(j, 1) = personalData(i, 3) And documentosPuestosData(j, 4) = proyectoCodigo Then
                        For k = 2 To UBound(masterDocumentosData, 1)
                            If masterDocumentosData(k, 1) = personalData(i, 3) And _
                               masterDocumentosData(k, 2) = documentosPuestosData(j, 2) Then
                                carpetasValidas.Add carpetaPersonal & "\" & masterDocumentosData(k, 3)
                                'ultimaFila = wsRutasValidas.Cells(wsRutasValidas.Rows.Count, 1).End(xlUp).Row + 1
                                'wsRutasValidas.Cells(ultimaFila, 1).Value = carpetaPersonal & "\" & masterDocumentosData(k, 3)
                                
                                carpetasValidas.Add carpetaRecursos & "\PERSONAL\" & documentoNombre
                                'ultimaFila = wsRutasValidas.Cells(wsRutasValidas.Rows.Count, 1).End(xlUp).Row + 1
                                'wsRutasValidas.Cells(ultimaFila, 1).Value = carpetaRecursos & "\PERSONAL\" & documentoNombre
                                
                                carpetasValidas.Add carpetaRecursos & "\PERSONAL\" & documentoNombre & "\" & masterDocumentosData(k, 3)
                                'ultimaFila = wsRutasValidas.Cells(wsRutasValidas.Rows.Count, 1).End(xlUp).Row + 1
                                'wsRutasValidas.Cells(ultimaFila, 1).Value = carpetaRecursos & "\PERSONAL\" & documentoNombre & "\" & masterDocumentosData(k, 3)
                            End If
                        Next k
                    End If
                Next j
            End If
        Next i

        ' Obtener carpetas válidas para Vehículos
        For i = 1 To UBound(vehiculosData, 1)
            If CStr(vehiculosData(i, 5)) = proyectoCodigo Then
                vehiculoNombre = vehiculosData(i, 1)
                carpetaVehiculos = carpetaHabilitaciones & "\VEHICULOS\" & vehiculoNombre
                carpetasValidas.Add carpetaVehiculos
                'ultimaFila = wsRutasValidas.Cells(wsRutasValidas.Rows.Count, 1).End(xlUp).Row + 1
                'wsRutasValidas.Cells(ultimaFila, 1).Value = carpetaVehiculos

                ' Buscar requisitos relacionados al vehículo
                For j = 2 To UBound(requisitosVehiculosData, 1)
                    If requisitosVehiculosData(j, 1) = vehiculosData(i, 2) And _
                       requisitosVehiculosData(j, 2) = vehiculosData(i, 3) And _
                       requisitosVehiculosData(j, 5) = proyectoCodigo Then
                        For k = 2 To UBound(masterRequisitosData, 1)
                            If masterRequisitosData(k, 1) = vehiculosData(i, 2) And _
                               masterRequisitosData(k, 2) = vehiculosData(i, 3) And _
                               masterRequisitosData(k, 3) = requisitosVehiculosData(j, 3) Then
                                carpetasValidas.Add carpetaVehiculos & "\" & masterRequisitosData(k, 4)
                                'ultimaFila = wsRutasValidas.Cells(wsRutasValidas.Rows.Count, 1).End(xlUp).Row + 1
                                'wsRutasValidas.Cells(ultimaFila, 1).Value = carpetaVehiculos & "\" & masterRequisitosData(k, 4)
                                
                                carpetasValidas.Add carpetaRecursos & "\VEHICULOS\" & vehiculoNombre
                                'ultimaFila = wsRutasValidas.Cells(wsRutasValidas.Rows.Count, 1).End(xlUp).Row + 1
                                'wsRutasValidas.Cells(ultimaFila, 1).Value = carpetaRecursos & "\VEHICULOS\" & vehiculoNombre
                                
                                carpetasValidas.Add carpetaRecursos & "\VEHICULOS\" & vehiculoNombre & "\" & masterRequisitosData(k, 4)
                                'ultimaFila = wsRutasValidas.Cells(wsRutasValidas.Rows.Count, 1).End(xlUp).Row + 1
                                'wsRutasValidas.Cells(ultimaFila, 1).Value = carpetaRecursos & "\VEHICULOS\" & vehiculoNombre & "\" & masterRequisitosData(k, 4)
                            End If
                        Next k
                    End If
                Next j
            End If
        Next i
    Next proyectoCodigo

    ' Después de haber recopilado todas las rutas válidas, proceder a eliminar las carpetas innecesarias
    Set fso = CreateObject("Scripting.FileSystemObject")

    For Each proyectoCodigo In proyectoCodigos
        carpetaProyecto = carpetaRaiz & proyectoCodigo
        carpetaHabilitaciones = carpetaProyecto & "\HABILITACIONES"

        ' Eliminar carpetas innecesarias en HABILITACIONES\PERSONAL
        If fso.FolderExists(carpetaHabilitaciones & "\PERSONAL") Then
            For Each carpeta In fso.GetFolder(carpetaHabilitaciones & "\PERSONAL").SubFolders
                If Not EsCarpetaValida(carpeta.Path, carpetasValidas) Then
                    fso.DeleteFolder carpeta.Path, True
                Else
                    ' Eliminar subcarpetas innecesarias dentro de las carpetas válidas de PERSONAL
                    For Each subcarpeta In carpeta.SubFolders
                        If Not EsCarpetaValida(subcarpeta.Path, carpetasValidas) Then
                            fso.DeleteFolder subcarpeta.Path, True
                        End If
                    Next subcarpeta
                End If
            Next carpeta
        End If

        ' Eliminar carpetas innecesarias en HABILITACIONES\VEHICULOS
        If fso.FolderExists(carpetaHabilitaciones & "\VEHICULOS") Then
            For Each carpeta In fso.GetFolder(carpetaHabilitaciones & "\VEHICULOS").SubFolders
                If Not EsCarpetaValida(carpeta.Path, carpetasValidas) Then
                    fso.DeleteFolder carpeta.Path, True
                Else
                    ' Eliminar subcarpetas innecesarias dentro de las carpetas válidas de VEHICULOS
                    For Each subcarpeta In carpeta.SubFolders
                        If Not EsCarpetaValida(subcarpeta.Path, carpetasValidas) Then
                            fso.DeleteFolder subcarpeta.Path, True
                        End If
                    Next subcarpeta
                End If
            Next carpeta
        End If

        ' Eliminar carpetas innecesarias en RECURSOS\PERSONAL
        If fso.FolderExists(carpetaRecursos & "\PERSONAL") Then
            For Each carpeta In fso.GetFolder(carpetaRecursos & "\PERSONAL").SubFolders
                If Not EsCarpetaValida(carpeta.Path, carpetasValidas) Then
                    fso.DeleteFolder carpeta.Path, True
                Else
                    ' Eliminar subcarpetas innecesarias dentro de las carpetas válidas de PERSONAL en RECURSOS
                    For Each subcarpeta In carpeta.SubFolders
                        If Not EsCarpetaValida(subcarpeta.Path, carpetasValidas) Then
                            fso.DeleteFolder subcarpeta.Path, True
                        End If
                    Next subcarpeta
                End If
            Next carpeta
        End If

        ' Eliminar carpetas innecesarias en RECURSOS\VEHICULOS
        If fso.FolderExists(carpetaRecursos & "\VEHICULOS") Then
            For Each carpeta In fso.GetFolder(carpetaRecursos & "\VEHICULOS").SubFolders
                If Not EsCarpetaValida(carpeta.Path, carpetasValidas) Then
                    fso.DeleteFolder carpeta.Path, True
                Else
                    ' Eliminar subcarpetas innecesarias dentro de las carpetas válidas de VEHICULOS en RECURSOS
                    For Each subcarpeta In carpeta.SubFolders
                        If Not EsCarpetaValida(subcarpeta.Path, carpetasValidas) Then
                            fso.DeleteFolder subcarpeta.Path, True
                        End If
                    Next subcarpeta
                End If
            Next carpeta
        End If
    Next proyectoCodigo

    If IsMissing(Logical) Or IsEmpty(Logical) Or Logical = "" Then
        MsgBox "Carpetas innecesarias eliminadas correctamente."
    End If
End Sub

Function EsCarpetaValida(carpetaPath As String, carpetasValidas As Collection) As Boolean
    Dim carpetaValida As Variant
    EsCarpetaValida = False
    For Each carpetaValida In carpetasValidas
        If carpetaValida = carpetaPath Then
            EsCarpetaValida = True
            Exit For
        End If
    Next carpetaValida
End Function


Sub CopiarInformacionSeleccionadaDesdeOtroLibro()
    Dim rutaArchivo As String
    Dim libroOrigen As Workbook
    Dim hoja As Worksheet
    Dim hojaDestino As Worksheet
    Dim nombreHoja As Variant
    Dim hojasCopiar As Variant
    Dim hojaParcial1 As String
    Dim hojaParcial2 As String
    
    ' Definir las hojas a copiar
    hojasCopiar = Array("MASTER PERSONAL", "MASTER VEHICULOS", "CATALOGOS", "DOCUMENTO X PUESTO", _
                        "REQUISITOS X VEHICULO", "PERSONAL", "VEHICULOS", "PERSONAL DOCUMENTOS", _
                        "VEHICULO DOCUMENTOS", "MASTER REQUISITOS VEHICULO", "MASTER DOCUMENTOS PERSONAL")
    hojaParcial1 = "MASTER REQUISITOS VEHICULO"
    hojaParcial2 = "MASTER DOCUMENTOS PERSONAL"
    
    ' Obtener la ruta del archivo desde la celda B8 de la hoja PRINCIPAL
    rutaArchivo = ThisWorkbook.Sheets("PRINCIPAL").Range("B8").Value
    
    ' Verificar si la ruta del archivo no está vacía
    If rutaArchivo = "" Then
        MsgBox "Por favor, especifique la ruta del archivo en la celda B8 de la hoja PRINCIPAL.", vbExclamation
        Exit Sub
    End If
    
    ' Abrir el libro origen
    On Error GoTo ErrorHandler
    Set libroOrigen = Workbooks.Open(rutaArchivo)
    On Error GoTo 0
    
    ' Recorrer las hojas especificadas
    For Each nombreHoja In hojasCopiar
        ' Verificar si la hoja existe en el libro origen
        On Error Resume Next
        Set hoja = libroOrigen.Sheets(nombreHoja)
        On Error GoTo 0
        
        If Not hoja Is Nothing Then
            ' Verificar si la hoja ya existe en el libro actual
            On Error Resume Next
            Set hojaDestino = ThisWorkbook.Sheets(nombreHoja)
            On Error GoTo 0
            
            ' Si la hoja existe, eliminarla
            If Not hojaDestino Is Nothing Then
                Application.DisplayAlerts = False
                hojaDestino.Delete
                Application.DisplayAlerts = True
                Set hojaDestino = Nothing
            End If
            
            ' Copiar la hoja del libro origen al libro actual
            If nombreHoja = hojaParcial1 Then
                ' Copiar solo hasta la columna D
                hoja.Range("A:D").Copy
                ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = nombreHoja
                ThisWorkbook.Sheets(nombreHoja).Range("A1").PasteSpecial Paste:=xlPasteAll
            ElseIf nombreHoja = hojaParcial2 Then
                ' Copiar solo hasta la columna C
                hoja.Range("A:C").Copy
                ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = nombreHoja
                ThisWorkbook.Sheets(nombreHoja).Range("A1").PasteSpecial Paste:=xlPasteAll
            Else
                hoja.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = nombreHoja
            End If
            
            ' Eliminar los nombres de rangos definidos en la hoja destino
            On Error Resume Next
            For Each n In ThisWorkbook.Names
                If InStr(1, n.RefersTo, "'" & nombreHoja & "'!") > 0 Then n.Delete
            Next n
            On Error GoTo 0
        End If
    Next nombreHoja
    
    ' Cerrar el libro origen sin guardar cambios
    libroOrigen.Close SaveChanges:=False
    
    MsgBox "La información ha sido copiada con éxito.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Hubo un error al abrir el libro origen. Verifique la ruta y el archivo.", vbCritical
End Sub


Sub ActualizarPersonalDocumentos(hojaOrigenNombre As String)
    Application.ScreenUpdating = False ' Desactivar la actualización de pantalla
    Application.DisplayAlerts = False ' Desactivar alertas
    Application.Calculation = xlCalculationManual ' Desactivar cálculos automáticos

    On Error GoTo ErrorHandler

    Dim wsPersonalDocs As Worksheet
    Dim wsOrigen As Worksheet
    Dim filaDestino As Range
    Dim DocIdentidad As String
    Dim documento As String
    Dim emision As Date
    Dim vigencia As Date
    Dim observacion As String
    Dim estado As Long
    Dim i As Long
    Dim lastRowOrigen As Long
    Dim lastRowDocs As Long
    Dim found As Boolean

    ' Definir la hoja PERSONAL DOCUMENTOS
    Set wsPersonalDocs = ThisWorkbook.Sheets("PERSONAL DOCUMENTOS")

    ' Buscar la hoja de origen por su nombre
    On Error Resume Next
    Set wsOrigen = ThisWorkbook.Sheets(hojaOrigenNombre)
    On Error GoTo 0

    ' Verificar si la hoja de origen existe
    If wsOrigen Is Nothing Then
        MsgBox "La hoja " & hojaOrigenNombre & " no existe.", vbExclamation
        Exit Sub
    End If

    ' Determinar la última fila con datos en la hoja de origen (asumiendo que la columna 1 siempre tiene datos)
    lastRowOrigen = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row

    ' Determinar la última fila con datos en la hoja PERSONAL DOCUMENTOS
    lastRowDocs = wsPersonalDocs.Cells(wsPersonalDocs.Rows.Count, 1).End(xlUp).Row

    ' Iterar a través de cada fila de la hoja de origen desde la fila 2 (asumiendo que la fila 1 tiene encabezados)
    For i = 2 To lastRowOrigen
        ' Verificar si el valor de la columna F (Columna 6) es igual a 0
        If wsOrigen.Cells(i, 6).Value = 0 Then
            ' Obtener datos de la fila de origen
            DocIdentidad = wsOrigen.Cells(i, 1).Value  ' Columna 1 tiene el DocIdentidad
            documento = wsOrigen.Cells(i, 2).Value     ' Columna 2 tiene el documento
            emision = wsOrigen.Cells(i, 3).Value       ' Columna 3 tiene la fecha de emisión
            vigencia = wsOrigen.Cells(i, 4).Value      ' Columna 4 tiene la fecha de vigencia
            observacion = wsOrigen.Cells(i, 5).Value   ' Columna 5 tiene las observaciones

            found = False
            ' Iterar a través de las filas de PERSONAL DOCUMENTOS para buscar coincidencias
            Dim j As Long
            For j = 2 To lastRowDocs
                If wsPersonalDocs.Cells(j, 1).Value = DocIdentidad And wsPersonalDocs.Cells(j, 2).Value = documento Then
                    ' Si se encuentra la coincidencia, actualizar los datos
                    wsPersonalDocs.Cells(j, 3).Value = emision
                    wsPersonalDocs.Cells(j, 4).Value = vigencia
                    wsPersonalDocs.Cells(j, 5).Value = observacion
                    found = True
                    Exit For
                End If
            Next j

            ' Si no se encontró la combinación, agregar una nueva fila en PERSONAL DOCUMENTOS
            If Not found Then
                lastRowDocs = lastRowDocs + 1
                wsPersonalDocs.Cells(lastRowDocs, 1).Value = DocIdentidad
                wsPersonalDocs.Cells(lastRowDocs, 2).Value = documento
                wsPersonalDocs.Cells(lastRowDocs, 3).Value = emision
                wsPersonalDocs.Cells(lastRowDocs, 4).Value = vigencia
                wsPersonalDocs.Cells(lastRowDocs, 5).Value = observacion
            End If

            ' Cambiar el valor de la columna F (estado) a 1 en la hoja de origen
            wsOrigen.Cells(i, 6).Value = 1
        End If
    Next i

Finalizar:
    ' Restaurar configuración original
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Exit Sub

ErrorHandler:
    MsgBox "Error al actualizar los datos: " & Err.Description, vbExclamation
    Resume Finalizar
End Sub

Function CopiarArchivoConNuevoNombre(archivoRuta As String, link As String, Abreviatura As String, Nombres As String, Fecha As Date) As Boolean
    Dim destinoArchivo As String
    Dim nuevoNombre As String
    Dim extension As String

    ' Obtener la extensión del archivo original
    extension = ObtenerExtension(archivoRuta)
    
    ' Formatear la fecha al formato que prefieras (por ejemplo, YYYY-MM-DD)
    Dim fechaFormateada As String
    fechaFormateada = Format(Fecha, "dd.mm.yyyy")
    
    ' Crear el nuevo nombre incluyendo la extensión original
    nuevoNombre = Nombres & " - " & Abreviatura & " " & fechaFormateada & extension
    
    ' Construir la ruta completa del archivo de destino
    destinoArchivo = link & "\" & nuevoNombre
    
    ' Copiar el archivo seleccionado al destino especificado
    On Error GoTo ErrorHandler
    FileCopy archivoRuta, destinoArchivo
    On Error GoTo 0
    
    ' Si la copia fue exitosa, devolvemos True
    CopiarArchivoConNuevoNombre = True
    Exit Function

ErrorHandler:
    ' En caso de error, devolver False
    MsgBox "Error al copiar el archivo: " & Err.Description
    CopiarArchivoConNuevoNombre = False
End Function

Function ObtenerExtension(archivoRuta As String) As String
    Dim extension As String
    Dim puntoPosicion As Integer

    ' Encontrar la posición del último punto (.) en la ruta del archivo
    puntoPosicion = InStrRev(archivoRuta, ".")

    ' Extraer la extensión desde la última posición del punto hasta el final del archivo
    If puntoPosicion > 0 Then
        extension = Right(archivoRuta, Len(archivoRuta) - puntoPosicion + 1)
    Else
        ' Si no se encuentra un punto, asumimos que no tiene extensión
        extension = ""
    End If

    ' Devolver la extensión del archivo
    ObtenerExtension = extension
End Function


Function ObtenerInicialesMayusculas(texto As String) As String
    Dim palabras() As String
    Dim iniciales As String
    Dim palabra As Variant
    
    ' Dividir el texto en palabras utilizando el espacio como delimitador
    palabras = Split(texto, " ")
    
    ' Inicializar la variable iniciales
    iniciales = ""
    
    ' Recorrer cada palabra y tomar la primera letra en mayúscula
    For Each palabra In palabras
        If Len(palabra) > 0 Then ' Verificar que la palabra no esté vacía
            iniciales = iniciales & UCase(Left(palabra, 1))
        End If
    Next palabra
    
    ' Devolver las iniciales
    ObtenerInicialesMayusculas = iniciales
End Function




