Attribute VB_Name = "Utilitarios"
Function UniqueValuesVertical(rng As Range) As Variant
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim cell As Range
    For Each cell In rng
        If Not dict.exists(cell.Value) Then
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
        If Not IsEmpty(cell.Value) And Not dict.exists(cell.Value) Then
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
            If Not dict.exists(datos(i, colResIndex)) Then
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
    Dim CarpetaBase As String
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

    CarpetaBase = wsPrincipal.Cells(6, 2).Value
    Debug.Print "Carpeta base: " & CarpetaBase
    
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
        If folder.files.Count > 0 Then
            ' Copiar todos los archivos de la carpeta de origen a la carpeta de destino
                For Each file In folder.files
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
    Dim CarpetaBase As String
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

    CarpetaBase = wsPrincipal.Cells(6, 2).Value
    Debug.Print "Carpeta base: " & CarpetaBase
    
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
        If folder.files.Count > 0 Then
            ' Copiar todos los archivos de la carpeta de origen a la carpeta de destino
                For Each file In folder.files
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
            If Not dictDocumentos.exists(documento) Then
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
            If Not dictTrabajadores.exists(trabajador) Then
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
                 If EstadoDoc = "" And observacion <> "APROBADO POR GERENCIA" And observacion <> "NO RESTRICTIVO" And observacion <> "NO APLICA" And observacion <> "OBSERVADO" And observacion <> "SUSTENTO MEDICO" Then
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
                                If observacion = "SUSTENTO MEDICO" Then
                                    wsGestionIntegral.Cells(rowIndex, colIndex).Value = "SM"
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
                                    If observacion = "SUSTENTO MEDICO" Then
                                        wsGestionIntegral.Cells(rowIndex, colIndex).Value = "SM"
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
                    End If
                Else
                    If (encontrado = False Or encontrado = Falso) And observacion <> "APROBADO POR GERENCIA" And observacion <> "NO RESTRICTIVO" And observacion <> "NO APLICA" And observacion <> "SUSTENTO MEDICO" Then
                        wsGestionIntegral.Cells(rowIndex, colIndex).Value = "FALTA"
                    Else
                    'Modificación
                        If observacion = "APROBADO POR GERENCIA" Then
                            wsGestionIntegral.Cells(rowIndex, colIndex).Value = "AGA"
                        Else
                            If observacion = "NO RESTRICTIVO" Then
                                wsGestionIntegral.Cells(rowIndex, colIndex).Value = "NR"
                            Else
                                If observacion = "SUSTENTO MEDICO" Then
                                        wsGestionIntegral.Cells(rowIndex, colIndex).Value = "SM"
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
            If Not dictDocumentos.exists(documento) Then
                dictDocumentos.Add documento, documento
            End If
        End If
    Next celda

    ' Obtener vehículos del proyecto de la tabla VehiculoProyecto
    lastRow = wsVehiculo.Cells(wsVehiculo.Rows.Count, "A").End(xlUp).Row
    For Each celda In wsVehiculo.Range("A2:A" & lastRow)
        If celda.Offset(0, 4).Value = proyectoCodigo And celda.Offset(0, 1).Value <> "EMBARCACIÓN" Then
            vehiculo = celda.Value
            If Not dictVehiculos.exists(vehiculo) Then
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
            If Not dictDocumentos.exists(documento) Then
                dictDocumentos.Add documento, documento
            End If
        End If
    Next celda

    ' Obtener vehículos del proyecto de la tabla VehiculoProyecto
    lastRow = wsVehiculo.Cells(wsVehiculo.Rows.Count, "A").End(xlUp).Row
    For Each celda In wsVehiculo.Range("A2:A" & lastRow)
        If celda.Offset(0, 4).Value = proyectoCodigo And celda.Offset(0, 1).Value = "EMBARCACIÓN" Then
            vehiculo = celda.Value
            If Not dictVehiculos.exists(vehiculo) Then
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
    Dim CarpetaBase As String
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

    CarpetaBase = wsPrincipal.Cells(6, 2).Value
    Debug.Print "Carpeta base: " & CarpetaBase
    
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
        If folder.files.Count > 0 Then
            ' Copiar todos los archivos de la carpeta de origen a la carpeta de destino
                For Each file In folder.files
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
    ' Registrar hora de inicio
    Dim horaInicio As Date
    horaInicio = Now()
    ThisWorkbook.Sheets("CATALOGOS").Range("N2").Value = Format(horaInicio, "dd/mm/yyyy hh:mm:ss")
    
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
    Dim subCarpeta As Object
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
                    For Each subCarpeta In carpeta.SubFolders
                        If Not EsCarpetaValida(subCarpeta.Path, carpetasValidas) Then
                            fso.DeleteFolder subCarpeta.Path, True
                        End If
                    Next subCarpeta
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
                    For Each subCarpeta In carpeta.SubFolders
                        If Not EsCarpetaValida(subCarpeta.Path, carpetasValidas) Then
                            fso.DeleteFolder subCarpeta.Path, True
                        End If
                    Next subCarpeta
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
                    For Each subCarpeta In carpeta.SubFolders
                        If Not EsCarpetaValida(subCarpeta.Path, carpetasValidas) Then
                            fso.DeleteFolder subCarpeta.Path, True
                        End If
                    Next subCarpeta
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
                    For Each subCarpeta In carpeta.SubFolders
                        If Not EsCarpetaValida(subCarpeta.Path, carpetasValidas) Then
                            fso.DeleteFolder subCarpeta.Path, True
                        End If
                    Next subCarpeta
                End If
            Next carpeta
        End If
    Next proyectoCodigo

    ' Registrar hora de finalización
    ThisWorkbook.Sheets("CATALOGOS").Range("N3").Value = Format(Now(), "dd/mm/yyyy hh:mm:ss")
    
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

Sub EliminarCarpetasInnecesariasMejorada(Optional Logical As Variant)
    ' Registrar hora de inicio
    Dim horaInicio As Date
    horaInicio = Now()
    ThisWorkbook.Sheets("CATALOGOS").Range("N2").Value = Format(horaInicio, "dd/mm/yyyy hh:mm:ss")
    
    ' Desactivar actualizaciones de pantalla para acelerar
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Dim wsPrincipal As Worksheet, wsPersonal As Worksheet, wsDocumentosPuestos As Worksheet
    Dim wsVehiculos As Worksheet, wsRequisitosVehiculos As Worksheet
    Dim wsMasterDocumentos As Worksheet, wsMasterRequisitos As Worksheet
    Dim proyectoCodigo As Variant, carpetaRaiz As String
    Dim carpetaProyecto As String, carpetaHabilitaciones As String, carpetaRecursos As String
    Dim carpetaPersonal As String, carpetaVehiculos As String
    Dim personalData As Variant, vehiculosData As Variant, documentosPuestosData As Variant
    Dim masterDocumentosData As Variant, requisitosVehiculosData As Variant, masterRequisitosData As Variant
    Dim i As Long, j As Long, k As Long
    Dim documentoNombre As String, vehiculoNombre As String
    Dim fso As Object
    Dim carpeta As Object, subCarpeta As Object
    
    ' Usar un diccionario para tener búsquedas O(1) en lugar de O(n) de una Collection
    Dim dictCarpetasValidas As Object
    Dim dictProyectos As Object
    
    Set dictCarpetasValidas = CreateObject("Scripting.Dictionary")
    Set dictProyectos = CreateObject("Scripting.Dictionary")
    
    ' Establecer las hojas
    Set wsPrincipal = ThisWorkbook.Sheets("PRINCIPAL")
    Set wsPersonal = ThisWorkbook.Sheets("PERSONAL")
    Set wsDocumentosPuestos = ThisWorkbook.Sheets("DOCUMENTO X PUESTO")
    Set wsVehiculos = ThisWorkbook.Sheets("VEHICULOS")
    Set wsRequisitosVehiculos = ThisWorkbook.Sheets("REQUISITOS X VEHICULO")
    Set wsMasterDocumentos = ThisWorkbook.Sheets("MASTER DOCUMENTOS PERSONAL")
    Set wsMasterRequisitos = ThisWorkbook.Sheets("MASTER REQUISITOS VEHICULO")

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

    ' Crear diccionarios para búsquedas más rápidas
    Dim dictDocsPuestos As Object, dictDocsMaster As Object
    Dim dictReqVehiculos As Object, dictReqMaster As Object
    
    Set dictDocsPuestos = CreateObject("Scripting.Dictionary")
    Set dictDocsMaster = CreateObject("Scripting.Dictionary")
    Set dictReqVehiculos = CreateObject("Scripting.Dictionary")
    Set dictReqMaster = CreateObject("Scripting.Dictionary")
    
    ' Preprocesar datos de puestos/documentos para búsquedas más rápidas
    For j = 2 To UBound(documentosPuestosData, 1)
        Dim puestoProyKey As String
        puestoProyKey = CStr(documentosPuestosData(j, 1)) & "|" & CStr(documentosPuestosData(j, 4))
        
        If Not dictDocsPuestos.exists(puestoProyKey) Then
            Dim docsCollection As New Collection
            dictDocsPuestos.Add puestoProyKey, docsCollection
        End If
        
        On Error Resume Next ' Prevenir duplicados
        dictDocsPuestos(puestoProyKey).Add CStr(documentosPuestosData(j, 2))
        On Error GoTo 0
    Next j
    
    ' Preprocesar datos de masterDocumentos para búsquedas más rápidas
    For k = 2 To UBound(masterDocumentosData, 1)
        Dim docMasterKey As String
        docMasterKey = CStr(masterDocumentosData(k, 1)) & "|" & CStr(masterDocumentosData(k, 2))
        
        If Not dictDocsMaster.exists(docMasterKey) Then
            dictDocsMaster.Add docMasterKey, CStr(masterDocumentosData(k, 3))
        End If
    Next k
    
    ' Preprocesar datos de requisitosVehiculos para búsquedas más rápidas
    For j = 2 To UBound(requisitosVehiculosData, 1)
        Dim reqVehKey As String
        reqVehKey = CStr(requisitosVehiculosData(j, 1)) & "|" & CStr(requisitosVehiculosData(j, 2)) & "|" & CStr(requisitosVehiculosData(j, 5))
        
        If Not dictReqVehiculos.exists(reqVehKey) Then
            Dim reqCollection As New Collection
            dictReqVehiculos.Add reqVehKey, reqCollection
        End If
        
        On Error Resume Next ' Prevenir duplicados
        dictReqVehiculos(reqVehKey).Add CStr(requisitosVehiculosData(j, 3))
        On Error GoTo 0
    Next j
    
    ' Preprocesar datos de masterRequisitos para búsquedas más rápidas
    For k = 2 To UBound(masterRequisitosData, 1)
        Dim reqMasterKey As String
        reqMasterKey = CStr(masterRequisitosData(k, 1)) & "|" & CStr(masterRequisitosData(k, 2)) & "|" & CStr(masterRequisitosData(k, 3))
        
        If Not dictReqMaster.exists(reqMasterKey) Then
            dictReqMaster.Add reqMasterKey, CStr(masterRequisitosData(k, 4))
        End If
    Next k

    ' Obtener códigos de proyecto únicos de la hoja PERSONAL
    For i = 1 To UBound(personalData, 1)
        If Not IsEmpty(personalData(i, 5)) Then
            Dim proyecto As String
            proyecto = CStr(personalData(i, 5))
            If Not dictProyectos.exists(proyecto) Then
                dictProyectos.Add proyecto, True
            End If
        End If
    Next i

    ' Obtener códigos de proyecto únicos de la hoja VEHICULOS
    For i = 1 To UBound(vehiculosData, 1)
        If Not IsEmpty(vehiculosData(i, 5)) Then
            proyecto = CStr(vehiculosData(i, 5))
            If Not dictProyectos.exists(proyecto) Then
                dictProyectos.Add proyecto, True
            End If
        End If
    Next i

    ' Loop para cada código de proyecto (Solo para recopilar rutas válidas)
    For Each proyectoCodigo In dictProyectos.keys
        carpetaProyecto = carpetaRaiz & proyectoCodigo
        carpetaHabilitaciones = carpetaProyecto & "\HABILITACIONES"
        
        ' Obtener carpetas válidas para Personal
        For i = 1 To UBound(personalData, 1)
            If CStr(personalData(i, 5)) = proyectoCodigo Then
                documentoNombre = Trim(CStr(personalData(i, 1)))
                carpetaPersonal = carpetaHabilitaciones & "\PERSONAL\" & documentoNombre
                
                ' Añadir a diccionario de carpetas válidas
                If Not dictCarpetasValidas.exists(carpetaPersonal) Then
                    dictCarpetasValidas.Add carpetaPersonal, True
                End If
                
                ' Buscar documentos relacionados al puesto
                Dim puesto As String, puestoProyectoKey As String
                puesto = CStr(personalData(i, 3))
                puestoProyectoKey = puesto & "|" & proyectoCodigo
                
                If dictDocsPuestos.exists(puestoProyectoKey) Then
                    Dim documento As Variant
                    For Each documento In dictDocsPuestos(puestoProyectoKey)
                        Dim docKey As String
                        docKey = puesto & "|" & documento
                        
                        If dictDocsMaster.exists(docKey) Then
                            Dim nombreDoc As String
                            nombreDoc = dictDocsMaster(docKey)
                            
                            ' Agregar carpetas válidas a diccionario
                            If Not dictCarpetasValidas.exists(carpetaPersonal & "\" & nombreDoc) Then
                                dictCarpetasValidas.Add carpetaPersonal & "\" & nombreDoc, True
                            End If
                            
                            If Not dictCarpetasValidas.exists(carpetaRecursos & "\PERSONAL\" & documentoNombre) Then
                                dictCarpetasValidas.Add carpetaRecursos & "\PERSONAL\" & documentoNombre, True
                            End If
                            
                            If Not dictCarpetasValidas.exists(carpetaRecursos & "\PERSONAL\" & documentoNombre & "\" & nombreDoc) Then
                                dictCarpetasValidas.Add carpetaRecursos & "\PERSONAL\" & documentoNombre & "\" & nombreDoc, True
                            End If
                        End If
                    Next documento
                End If
            End If
        Next i

        ' Obtener carpetas válidas para Vehículos
        For i = 1 To UBound(vehiculosData, 1)
            If CStr(vehiculosData(i, 5)) = proyectoCodigo Then
                vehiculoNombre = CStr(vehiculosData(i, 1))
                carpetaVehiculos = carpetaHabilitaciones & "\VEHICULOS\" & vehiculoNombre
                
                ' Añadir a diccionario de carpetas válidas
                If Not dictCarpetasValidas.exists(carpetaVehiculos) Then
                    dictCarpetasValidas.Add carpetaVehiculos, True
                End If

                ' Buscar requisitos relacionados al vehículo usando diccionarios
                Dim tipoVeh As String, categoriaVeh As String, vehKey As String
                tipoVeh = CStr(vehiculosData(i, 2))
                categoriaVeh = CStr(vehiculosData(i, 3))
                vehKey = tipoVeh & "|" & categoriaVeh & "|" & proyectoCodigo
                
                If dictReqVehiculos.exists(vehKey) Then
                    Dim requisito As Variant
                    For Each requisito In dictReqVehiculos(vehKey)
                        Dim reqKey As String
                        reqKey = tipoVeh & "|" & categoriaVeh & "|" & requisito
                        
                        If dictReqMaster.exists(reqKey) Then
                            Dim nombreReq As String
                            nombreReq = dictReqMaster(reqKey)
                            
                            ' Agregar carpetas válidas a diccionario
                            If Not dictCarpetasValidas.exists(carpetaVehiculos & "\" & nombreReq) Then
                                dictCarpetasValidas.Add carpetaVehiculos & "\" & nombreReq, True
                            End If
                            
                            If Not dictCarpetasValidas.exists(carpetaRecursos & "\VEHICULOS\" & vehiculoNombre) Then
                                dictCarpetasValidas.Add carpetaRecursos & "\VEHICULOS\" & vehiculoNombre, True
                            End If
                            
                            If Not dictCarpetasValidas.exists(carpetaRecursos & "\VEHICULOS\" & vehiculoNombre & "\" & nombreReq) Then
                                dictCarpetasValidas.Add carpetaRecursos & "\VEHICULOS\" & vehiculoNombre & "\" & nombreReq, True
                            End If
                        End If
                    Next requisito
                End If
            End If
        Next i
    Next proyectoCodigo

    ' Después de haber recopilado todas las rutas válidas, proceder a eliminar las carpetas innecesarias
    Set fso = CreateObject("Scripting.FileSystemObject")

    For Each proyectoCodigo In dictProyectos.keys
        carpetaProyecto = carpetaRaiz & proyectoCodigo
        carpetaHabilitaciones = carpetaProyecto & "\HABILITACIONES"
        
        ' Procesar todas las carpetas en paralelo en lugar de secuencial
        Dim carpetasARevisar As Object
        Set carpetasARevisar = CreateObject("Scripting.Dictionary")
        
        ' Añadir todas las rutas a revisar
        If fso.FolderExists(carpetaHabilitaciones & "\PERSONAL") Then
            carpetasARevisar.Add carpetaHabilitaciones & "\PERSONAL", "HABILITACIONES\PERSONAL"
        End If
        
        If fso.FolderExists(carpetaHabilitaciones & "\VEHICULOS") Then
            carpetasARevisar.Add carpetaHabilitaciones & "\VEHICULOS", "HABILITACIONES\VEHICULOS"
        End If
        
        If fso.FolderExists(carpetaRecursos & "\PERSONAL") Then
            carpetasARevisar.Add carpetaRecursos & "\PERSONAL", "RECURSOS\PERSONAL"
        End If
        
        If fso.FolderExists(carpetaRecursos & "\VEHICULOS") Then
            carpetasARevisar.Add carpetaRecursos & "\VEHICULOS", "RECURSOS\VEHICULOS"
        End If
        
        ' Procesar todas las carpetas de una sola vez para minimizar el cambio de contexto
        Dim rutaCarpeta As Variant
        For Each rutaCarpeta In carpetasARevisar.keys
            For Each carpeta In fso.GetFolder(rutaCarpeta).SubFolders
                If Not dictCarpetasValidas.exists(carpeta.Path) Then
                    fso.DeleteFolder carpeta.Path, True
                Else
                    ' Eliminar subcarpetas innecesarias dentro de las carpetas válidas
                    For Each subCarpeta In carpeta.SubFolders
                        If Not dictCarpetasValidas.exists(subCarpeta.Path) Then
                            fso.DeleteFolder subCarpeta.Path, True
                        End If
                    Next subCarpeta
                End If
            Next carpeta
        Next rutaCarpeta
    Next proyectoCodigo

    ' Restaurar configuración de Excel
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Registrar hora de finalización
    ThisWorkbook.Sheets("CATALOGOS").Range("N3").Value = Format(Now(), "dd/mm/yyyy hh:mm:ss")
    
    If IsMissing(Logical) Or IsEmpty(Logical) Or Logical = "" Then
        MsgBox "Carpetas innecesarias eliminadas correctamente."
    End If
End Sub


Sub EliminarCarpetasInnecesariasVersion2(Optional Logical As Variant)
    ' Registrar hora de inicio
    Dim horaInicio As Date
    horaInicio = Now()
    ThisWorkbook.Sheets("CATALOGOS").Range("N2").Value = Format(horaInicio, "dd/mm/yyyy hh:mm:ss")

    Dim wsPrincipal As Worksheet, wsPersonal As Worksheet, wsDocumentosPuestos As Worksheet
    Dim wsVehiculos As Worksheet, wsRequisitosVehiculos As Worksheet
    Dim wsMasterDocumentos As Worksheet, wsMasterRequisitos As Worksheet
    Dim carpetaRaiz As String, carpetaRecursos As String
    Dim personalData As Variant, vehiculosData As Variant
    Dim documentosPuestosData As Variant, requisitosVehiculosData As Variant
    Dim masterDocumentosData As Variant, masterRequisitosData As Variant
    Dim i As Long, j As Long, k As Long
    Dim documentoNombre As String, vehiculoNombre As String
    Dim fso As Object, carpeta As Object, subCarpeta As Object
    Dim carpetasValidas As Object
    Dim proyectoCodigos As Object
    Dim proyectoCodigo As Variant

    ' Establecer las hojas
    Set wsPrincipal = ThisWorkbook.Sheets("PRINCIPAL")
    Set wsPersonal = ThisWorkbook.Sheets("PERSONAL")
    Set wsDocumentosPuestos = ThisWorkbook.Sheets("DOCUMENTO X PUESTO")
    Set wsVehiculos = ThisWorkbook.Sheets("VEHICULOS")
    Set wsRequisitosVehiculos = ThisWorkbook.Sheets("REQUISITOS X VEHICULO")
    Set wsMasterDocumentos = ThisWorkbook.Sheets("MASTER DOCUMENTOS PERSONAL")
    Set wsMasterRequisitos = ThisWorkbook.Sheets("MASTER REQUISITOS VEHICULO")

    ' Obtener la carpeta raíz y asegurar que termina en "\"
    carpetaRaiz = wsPrincipal.Range("B6").Value
    If Right(carpetaRaiz, 1) <> "\" Then carpetaRaiz = carpetaRaiz & "\"
    carpetaRecursos = carpetaRaiz & "RECURSOS"

    ' Cargar los datos en arrays (evita acceso constante a celdas de Excel)
    personalData = wsPersonal.Range("A2:E" & wsPersonal.Cells(wsPersonal.Rows.Count, "A").End(xlUp).Row).Value
    vehiculosData = wsVehiculos.Range("A2:E" & wsVehiculos.Cells(wsVehiculos.Rows.Count, "A").End(xlUp).Row).Value
    documentosPuestosData = wsDocumentosPuestos.UsedRange.Value
    requisitosVehiculosData = wsRequisitosVehiculos.UsedRange.Value
    masterDocumentosData = wsMasterDocumentos.UsedRange.Value
    masterRequisitosData = wsMasterRequisitos.UsedRange.Value

    ' Crear diccionario para almacenar códigos de proyecto únicos
    Set proyectoCodigos = CreateObject("Scripting.Dictionary")

    ' Agregar códigos de proyecto únicos desde PERSONAL y VEHICULOS
    For i = 1 To UBound(personalData, 1)
        If Not proyectoCodigos.exists(personalData(i, 5)) Then proyectoCodigos.Add personalData(i, 5), True
    Next i
    For i = 1 To UBound(vehiculosData, 1)
        If Not proyectoCodigos.exists(vehiculosData(i, 5)) Then proyectoCodigos.Add vehiculosData(i, 5), True
    Next i

    ' Crear diccionario para almacenar carpetas válidas
    Set carpetasValidas = CreateObject("Scripting.Dictionary")

    ' Recopilar carpetas válidas para cada código de proyecto
    For Each proyectoCodigo In proyectoCodigos.keys
        Dim carpetaProyecto As String, carpetaHabilitaciones As String
        carpetaProyecto = carpetaRaiz & proyectoCodigo
        carpetaHabilitaciones = carpetaProyecto & "\HABILITACIONES"

        ' Personal
        For i = 1 To UBound(personalData, 1)
            If CStr(personalData(i, 5)) = proyectoCodigo Then
                documentoNombre = Trim(personalData(i, 1))
                Dim carpetaPersonal As String
                carpetaPersonal = carpetaHabilitaciones & "\PERSONAL\" & documentoNombre
                carpetasValidas(carpetaPersonal) = True
                carpetasValidas(carpetaRecursos & "\PERSONAL\" & documentoNombre) = True

                ' Documentos personales
                For j = 2 To UBound(documentosPuestosData, 1)
                    If documentosPuestosData(j, 1) = personalData(i, 3) And documentosPuestosData(j, 4) = proyectoCodigo Then
                        For k = 2 To UBound(masterDocumentosData, 1)
                            If masterDocumentosData(k, 1) = personalData(i, 3) And masterDocumentosData(k, 2) = documentosPuestosData(j, 2) Then
                                carpetasValidas(carpetaPersonal & "\" & masterDocumentosData(k, 3)) = True
                                carpetasValidas(carpetaRecursos & "\PERSONAL\" & documentoNombre & "\" & masterDocumentosData(k, 3)) = True
                            End If
                        Next k
                    End If
                Next j
            End If
        Next i

        ' Vehículos
        For i = 1 To UBound(vehiculosData, 1)
            If CStr(vehiculosData(i, 5)) = proyectoCodigo Then
                vehiculoNombre = vehiculosData(i, 1)
                Dim carpetaVehiculos As String
                carpetaVehiculos = carpetaHabilitaciones & "\VEHICULOS\" & vehiculoNombre
                carpetasValidas(carpetaVehiculos) = True
                carpetasValidas(carpetaRecursos & "\VEHICULOS\" & vehiculoNombre) = True

                ' Requisitos de vehículos
                For j = 2 To UBound(requisitosVehiculosData, 1)
                    If requisitosVehiculosData(j, 5) = proyectoCodigo Then
                        For k = 2 To UBound(masterRequisitosData, 1)
                            If masterRequisitosData(k, 1) = vehiculosData(i, 2) And masterRequisitosData(k, 2) = vehiculosData(i, 3) And masterRequisitosData(k, 3) = requisitosVehiculosData(j, 3) Then
                                carpetasValidas(carpetaVehiculos & "\" & masterRequisitosData(k, 4)) = True
                                carpetasValidas(carpetaRecursos & "\VEHICULOS\" & vehiculoNombre & "\" & masterRequisitosData(k, 4)) = True
                            End If
                        Next k
                    End If
                Next j
            End If
        Next i
    Next proyectoCodigo

    ' Eliminar carpetas innecesarias
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim carpetaPadre As Variant
    For Each carpetaPadre In Array(carpetaRecursos & "\PERSONAL", carpetaRecursos & "\VEHICULOS")
        If fso.FolderExists(carpetaPadre) Then
            For Each carpeta In fso.GetFolder(carpetaPadre).SubFolders
                If Not carpetasValidas.exists(carpeta.Path) Then fso.DeleteFolder carpeta.Path, True
            Next carpeta
        End If
    Next carpetaPadre

    ' Registrar hora de finalización
    ThisWorkbook.Sheets("CATALOGOS").Range("N3").Value = Format(Now(), "dd/mm/yyyy hh:mm:ss")

    If IsMissing(Logical) Or IsEmpty(Logical) Then MsgBox "Carpetas innecesarias eliminadas correctamente."
End Sub





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




