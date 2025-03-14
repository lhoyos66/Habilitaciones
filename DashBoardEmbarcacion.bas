Attribute VB_Name = "DashBoardEmbarcacion"
Sub EjecutarGenerarDashboardEmbarcacion()
    Dim wsDashboard As Worksheet
    Dim Proyecto As String
    Dim Empresa As String
    Dim Servicio As String
    Dim placa As String
    
    Set wsDashboard = ThisWorkbook.Sheets("DASHBOARD EMBARCACIONES")
    
    ' Obtener valores de las celdas específicas
    Proyecto = wsDashboard.Range("C2").Value
    Empresa = wsDashboard.Range("C4").Value
    Servicio = wsDashboard.Range("C6").Value
    placa = wsDashboard.Range("C8").Value
    
    ' Verificar si el proyecto ha sido ingresado
    If Proyecto = "" Then
        MsgBox "El proyecto es un campo obligatorio.", vbExclamation
        Exit Sub
    End If
    
    ' Llamar a la macro principal con los parámetros obtenidos
    Call GenerarDashboardEmbarcacion(Proyecto, 1, Empresa, Servicio, placa)
End Sub

Sub EjecutarEmbarcacionResumen()
    Dim wsDashboard As Worksheet
    Dim Proyecto As String
    Dim Empresa As String
    Dim ServicioOPuesto As String
    Dim placa As String
    
    Set wsDashboard = ThisWorkbook.Sheets("EMBARCACION RESUMEN")
    
    ' Obtener valores de las celdas específicas
    Proyecto = wsDashboard.Range("C2").Value
    Empresa = wsDashboard.Range("C4").Value
    ServicioOPuesto = wsDashboard.Range("C6").Value
    placa = wsDashboard.Range("C8").Value
    
    ' Verificar si el proyecto ha sido ingresado
    If Proyecto = "" Then
        MsgBox "El proyecto es un campo obligatorio.", vbExclamation
        Exit Sub
    End If
    
    ' Llamar a la macro principal con los parámetros obtenidos
    Call GenerarResumenEmbarcacion(Proyecto, Empresa, ServicioOPuesto, placa)
End Sub

Sub GenerarResumenEmbarcacion(Proyecto As String, Optional Empresa As String = "", Optional ServicioOPuesto As String = "", Optional placa As String = "")
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsDashboard As Worksheet
    Dim wsPrincipal As Worksheet
    Dim wsVehiculo As Worksheet
    Dim wsDocXServicio As Worksheet
    Dim wsVehiculoDocs As Worksheet
    Dim wsDocPorTipo As Worksheet
    Dim wsVehiculoResumen As Worksheet
    Dim diasMargen As Integer
    Dim protocolo As String 'Modificación

    Set wsDashboard = ThisWorkbook.Sheets("DASHBOARD EMBARCACIONES")
    Set wsPrincipal = ThisWorkbook.Sheets("PRINCIPAL")
    Set wsVehiculo = ThisWorkbook.Sheets("VEHICULOS")
    Set wsDocXServicio = ThisWorkbook.Sheets("MASTER REQUISITOS VEHICULO")
    Set wsVehiculoDocs = ThisWorkbook.Sheets("VEHICULO DOCUMENTOS")
    Set wsDocPorTipo = ThisWorkbook.Sheets("REQUISITOS X VEHICULO")
    Set wsVehiculoResumen = ThisWorkbook.Sheets("EMBARCACION RESUMEN")

    ' Limpiar la tabla PersonalResumen
    On Error Resume Next
    wsVehiculoResumen.ListObjects("EmbarcacionResumen").DataBodyRange.Delete
    On Error GoTo 0

    Dim rutaBase As String
    rutaBase = wsPrincipal.Range("B6").Value

    Dim rutaProyecto As String
    Dim rutaRecurso As String
    rutaProyecto = rutaBase & "\" & wsDashboard.Range("C2").Value & "\HABILITACIONES\VEHICULOS"
    rutaRecurso = rutaBase & "\RECURSOS\VEHICULOS"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim carpetaProyecto As Object
    Set carpetaProyecto = fso.GetFolder(rutaProyecto)

    Dim subcarpeta As Object
    Dim subSubCarpeta As Object

    Dim dniPersonal As String
    Dim fila As ListRow

    ' Cargar datos en arrays para mejorar el rendimiento
    Dim vehiculoData As Variant
    Dim docXServicioData As Variant
    Dim vehiculoDocsData As Variant
    Dim docPorTipoData As Variant

    vehiculoData = wsVehiculo.Range("A2:F" & wsVehiculo.Cells(wsVehiculo.Rows.Count, "A").End(xlUp).Row).Value
    docXServicioData = wsDocXServicio.Range("A2:E" & wsDocXServicio.Cells(wsDocXServicio.Rows.Count, "A").End(xlUp).Row).Value
    vehiculoDocsData = wsVehiculoDocs.Range("A2:E" & wsVehiculoDocs.Cells(wsVehiculoDocs.Rows.Count, "A").End(xlUp).Row).Value
    docPorTipoData = wsDocPorTipo.Range("A2:E" & wsDocPorTipo.Cells(wsDocPorTipo.Rows.Count, "A").End(xlUp).Row).Value

    Dim vehiculoDict As Object
    Set vehiculoDict = CreateObject("Scripting.Dictionary")
    Dim docXServicioDict As Object
    Set docXServicioDict = CreateObject("Scripting.Dictionary")
    Dim docXServicioDictEmp As Object 'Modificacion
    Set docXServicioDictEmp = CreateObject("Scripting.Dictionary") 'Modificacion
    Dim vehiculoDocsDict As Object
    Set vehiculoDocsDict = CreateObject("Scripting.Dictionary")
    Dim docPorTipoDict As Object
    Set docPorTipoDict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = LBound(vehiculoData, 1) To UBound(vehiculoData, 1)
        vehiculoDict(vehiculoData(i, 1)) = Array(vehiculoData(i, 2), vehiculoData(i, 3), vehiculoData(i, 4), vehiculoData(i, 5), vehiculoData(i, 6))
    Next i

    For i = LBound(docXServicioData, 1) To UBound(docXServicioData, 1)
        docXServicioDict(docXServicioData(i, 4) & "-" & docXServicioData(i, 1) & "-" & docXServicioData(i, 2)) = docXServicioData(i, 3)
    Next i
    
    'Modificacion
    For i = LBound(docXServicioData, 1) To UBound(docXServicioData, 1)
        docXServicioDictEmp(docXServicioData(i, 4) & "-" & docXServicioData(i, 1) & "-" & docXServicioData(i, 2)) = Array(docXServicioData(i, 2), docXServicioData(i, 5))
    Next i

    For i = LBound(vehiculoDocsData, 1) To UBound(vehiculoDocsData, 1)
        vehiculoDocsDict(vehiculoDocsData(i, 1) & "-" & vehiculoDocsData(i, 2)) = Array(vehiculoDocsData(i, 3), vehiculoDocsData(i, 4), vehiculoDocsData(i, 5))
    Next i

    For i = LBound(docPorTipoData, 1) To UBound(docPorTipoData, 1)
        docPorTipoDict(docPorTipoData(i, 5) & "-" & docPorTipoData(i, 1) & "-" & docPorTipoData(i, 2) & "-" & docPorTipoData(i, 3)) = docPorTipoData(i, 4)
    Next i

    diasMargen = wsDashboard.Range("G2").Value

    Dim updateCounter As Long
    updateCounter = 0

    Dim habilitados, nohabilitados, pendientes, obligatorios As Long
    Dim totalDocs As Long

    For Each subcarpeta In carpetaProyecto.SubFolders
        dniPersonal = subcarpeta.Name
        
        ' Filtrar por DNI si se ha especificado
        If placa <> "" And placa <> dniPersonal Then
            GoTo NextPLACA
        End If

        habilitados = 0
        habilitadosCli = 0 'Modificacion
        pendientes = 0
        nohabilitados = 0
        totalDocs = 0
        totalDocsCli = 0 'Modificacion
        obligatorios = 0
        ' Buscar información del personal
        If vehiculoDict.Exists(dniPersonal) Then
            Dim vehiculoInfo As Variant
            vehiculoInfo = vehiculoDict(dniPersonal)

            ' Filtrar por Empresa y ServicioOPuesto si se han especificado
            If (Empresa <> "" And Empresa <> vehiculoInfo(2)) Or (Servicio <> "" And Servicio <> vehiculoInfo(0)) Or (vehiculoInfo(0) <> "EMBARCACIÓN") Then
                    GoTo NextPLACA
            End If

            For Each subSubCarpeta In subcarpeta.SubFolders
                

                ' Buscar información del documento
                Dim nombreCarpeta As String
                nombreCarpeta = subSubCarpeta.Name
                Dim keyDoc As String
                keyDoc = nombreCarpeta & "-" & vehiculoInfo(0) & "-" & vehiculoInfo(1)
                If docXServicioDict.Exists(keyDoc) Then
                    Dim keyObligatorio As String
                    keyObligatorio = Proyecto & "-" & vehiculoInfo(0) & "-" & vehiculoInfo(1) & "-" & docXServicioDict(keyDoc)
                    docObligatorio = docPorTipoDict(keyObligatorio)
                    protocolo = UCase(docXServicioDictEmp(keyDoc)(1)) 'Modificacion
                    If docObligatorio = "SI" Then
                        ' Contar documentos habilitados
                        If protocolo = "TEMA" Then
                            totalDocs = totalDocs + 1
                        Else
                            totalDocsCli = totalDocsCli + 1
                        End If
                    End If


                    ' Buscar fechas de emisión y vigencia iterando
                    Dim keyVehiculoDocs As String
                    keyVehiculoDocs = dniPersonal & "-" & docXServicioDict(keyDoc)
                    If vehiculoDocsDict.Exists(keyVehiculoDocs) Then
                        Dim docsInfo As Variant
                        docsInfo = vehiculoDocsDict(keyVehiculoDocs)
                        Dim vigencia As Variant
                        vigencia = docsInfo(1)
                        observacion = docsInfo(2)

                        ' Calcular vigencia en días
                        Dim vigenciaDias As Long
                        vigenciaDias = DateDiff("d", Date, vigencia)
                    
                                                
                        If docObligatorio = "SI" Then
                        ' Contar documentos habilitados
                            obligatorios = obligatorios + 1
                            If vigenciaDias >= 0 Then
                                'Modificacion
                                If protocolo = "TEMA" Then
                                    If observacion <> "OBSERVADO" Then
                                        habilitados = habilitados + 1
                                    Else
                                        nohabilitados = nohabilitados + 1
                                    End If
                                Else
                                    If observacion <> "OBSERVADO" Then
                                        habilitadosCli = habilitadosCli + 1
                                    Else
                                        nohabilitados = nohabilitados + 1
                                    End If
                                End If
                            Else
                                If observacion = "APROBADO POR GERENCIA" Or observacion = "NO RESTRICTIVO" Or observacion = "NO APLICA" Then
                                    If protocolo = "TEMA" Then
                                        habilitados = habilitados + 1
                                    Else
                                        habilitadosCli = habilitadosCli + 1
                                    End If
                                Else
                                    If (Not IsEmpty(vigencia) And IsDate(vigencia)) Or observacion = "OBSERVADO" Then
                                        nohabilitados = nohabilitados + 1
                                    Else
                                        pendientes = pendientes + 1
                                    End If
                                End If
                            End If
                        End If

                    End If
                End If
            Next subSubCarpeta

            ' Crear una nueva fila en la tabla PersonalResumen
            Set fila = wsVehiculoResumen.ListObjects("EmbarcacionResumen").ListRows.Add
            fila.Range(1, 1).Value = Proyecto
            fila.Range(1, 2).Value = vehiculoInfo(2) ' Empresa
            fila.Range(1, 3).Value = vehiculoInfo(0) ' Servicio
            fila.Range(1, 4).Value = "'" & dniPersonal ' Doc Identidad
            fila.Range(1, 5).Value = vehiculoInfo(1) ' Tipo
            fila.Range(1, 6).Value = habilitados ' Documentos Habilitados
            fila.Range(1, 7).Value = totalDocs ' Total Documentos
            fila.Range(1, 8).Value = habilitadosCli ' Documentos Habilitados Cliente
            fila.Range(1, 9).Value = totalDocsCli ' Total Documentos Cliente
            If habilitados + habilitadosCli >= totalDocs + totalDocsCli Then 'Modificacion
                fila.Range(1, 10).Value = "Habilitado"
            Else
                If nohabilitados <> obligatorios Then
                    fila.Range(1, 10).Value = "No habilitado" ' No Habilitado
                Else
                    fila.Range(1, 10).Value = "Pendiente" ' Pendiente
                End If
            End If
            fila.Range(1, 11).Value = vehiculoInfo(4) ' Situación
        End If

NextPLACA:
    Next subcarpeta

    ' Restaurar el cálculo automático y la actualización de pantalla
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' Notificar al usuario
    MsgBox "Resumen Vehiculo generado correctamente", vbInformation
End Sub

Sub ActivaResumenEmbarcacion()
    ' Activar la hoja "EMBARCACION RESUMEN"
    ThisWorkbook.Sheets("EMBARCACION RESUMEN").Activate
End Sub

Sub DashBoardEmbarcacionDesdeResumen()
    Dim wsVehiculoResumen As Worksheet
    Dim wsDashboard As Worksheet
    Dim tblVehiculoResumen As ListObject
    Dim filaActiva As ListRow
    Dim Proyecto As String
    Dim Empresa As String
    Dim Servicio As String
    Dim placa As String

    Set wsVehiculoResumen = ThisWorkbook.Sheets("EMBARCACION RESUMEN")
    Set wsDashboard = ThisWorkbook.Sheets("DASHBOARD EMBARCACIONES")
    
    Set tblVehiculoResumen = wsVehiculoResumen.ListObjects("EmbarcacionResumen")
 
    ' Verificar si la celda activa está dentro de la tabla VehiculoResumen
    If Not Intersect(ActiveCell, tblVehiculoResumen.DataBodyRange) Is Nothing Then
        ' Obtener la fila activa de la tabla
        Set filaActiva = tblVehiculoResumen.ListRows(ActiveCell.Row - tblVehiculoResumen.HeaderRowRange.Row)

        ' Capturar los valores de la fila activa
        Proyecto = filaActiva.Range(1, 1).Value
        Empresa = filaActiva.Range(1, 2).Value
        Servicio = filaActiva.Range(1, 3).Value
        placa = filaActiva.Range(1, 4).Value

        ' Copiar los valores a las celdas correspondientes en la hoja DASHBOARD PERSONAL
        wsDashboard.Range("C2").Value = Proyecto
        wsDashboard.Range("C4").Value = Empresa
        wsDashboard.Range("C6").Value = Servicio
        wsDashboard.Range("C8").Value = placa

        ' Activar la hoja DASHBOARD PERSONAL
        wsDashboard.Activate

        ' Ejecutar la macro GenerarDashboardPersonal
        Call GenerarDashboardEmbarcacion(Proyecto, 2, Empresa, Servicio, placa)
    Else
        MsgBox "Por favor, seleccione una celda dentro de la tabla PersonalResumen.", vbExclamation
    End If
End Sub


Sub GenerarDashboardEmbarcacion(Proyecto As String, TipoRep As Long, Optional Empresa As String = "", Optional Servicio As String = "", Optional placa As String = "")
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsDashboard As Worksheet
    Dim wsPrincipal As Worksheet
    Dim wsVehiculo As Worksheet
    Dim wsDocXServicio As Worksheet
    Dim wsVehiculoDocs As Worksheet
    Dim wsDocPorServicio As Worksheet
    Dim diasMargen As Integer
    Dim link As String
    Dim cell As Range

    Set wsDashboard = ThisWorkbook.Sheets("DASHBOARD EMBARCACIONES")
    Set wsPrincipal = ThisWorkbook.Sheets("PRINCIPAL")
    Set wsVehiculo = ThisWorkbook.Sheets("VEHICULOS")
    Set wsDocXServicio = ThisWorkbook.Sheets("MASTER REQUISITOS VEHICULO")
    Set wsVehiculoDocs = ThisWorkbook.Sheets("VEHICULO DOCUMENTOS")
    Set wsDocPorServicio = ThisWorkbook.Sheets("REQUISITOS X VEHICULO")

    ' Limpiar la tabla DashboardEmbarcacion
    On Error Resume Next
    wsDashboard.ListObjects("DashboardEmbarcacion").DataBodyRange.Delete
    On Error GoTo 0
    
    Dim rutaBase As String
    rutaBase = wsPrincipal.Range("B6").Value

    Dim rutaProyecto As String
    Dim rutaRecurso As String
    rutaProyecto = rutaBase & "\" & wsDashboard.Range("C2").Value & "\HABILITACIONES\VEHICULOS"
    rutaRecurso = rutaBase & "\RECURSOS\VEHICULOS"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim carpetaProyecto As Object
    Set carpetaProyecto = fso.GetFolder(rutaProyecto)

    Dim subcarpeta As Object
    Dim subSubCarpeta As Object

    Dim placaVehiculo As String
    Dim fila As ListRow

    ' Cargar datos en arrays para mejorar el rendimiento
    Dim vehiculoData As Variant
    Dim docXServicioData As Variant
    Dim vehiculoDocsData As Variant
    Dim docPorServicioData As Variant

    vehiculoData = wsVehiculo.Range("A2:E" & wsVehiculo.Cells(wsVehiculo.Rows.Count, "A").End(xlUp).Row).Value
    docXServicioData = wsDocXServicio.Range("A2:D" & wsDocXServicio.Cells(wsDocXServicio.Rows.Count, "A").End(xlUp).Row).Value
    vehiculoDocsData = wsVehiculoDocs.Range("A2:F" & wsVehiculoDocs.Cells(wsVehiculoDocs.Rows.Count, "A").End(xlUp).Row).Value
    docPorServicioData = wsDocPorServicio.Range("A2:E" & wsDocPorServicio.Cells(wsDocPorServicio.Rows.Count, "A").End(xlUp).Row).Value

    Dim vehiculoDict As Object
    Set vehiculoDict = CreateObject("Scripting.Dictionary")
    Dim docXServicioDict As Object
    Set docXServicioDict = CreateObject("Scripting.Dictionary")
    Dim vehiculoDocsDict As Object
    Set vehiculoDocsDict = CreateObject("Scripting.Dictionary")
    Dim docPorServicioDict As Object
    Set docPorServicioDict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = LBound(vehiculoData, 1) To UBound(vehiculoData, 1)
        vehiculoDict(vehiculoData(i, 1)) = Array(vehiculoData(i, 2), vehiculoData(i, 3), vehiculoData(i, 4), vehiculoData(i, 5))
    Next i

    For i = LBound(docXServicioData, 1) To UBound(docXServicioData, 1)
        docXServicioDict(docXServicioData(i, 4) & "-" & docXServicioData(i, 1) & "-" & docXServicioData(i, 2)) = docXServicioData(i, 3)
    Next i

    For i = LBound(vehiculoDocsData, 1) To UBound(vehiculoDocsData, 1)
        vehiculoDocsDict(vehiculoDocsData(i, 1) & "-" & vehiculoDocsData(i, 2)) = Array(vehiculoDocsData(i, 3), vehiculoDocsData(i, 4), vehiculoDocsData(i, 5))
    Next i

    For i = LBound(docPorServicioData, 1) To UBound(docPorServicioData, 1)
        docPorServicioDict(docPorServicioData(i, 5) & "-" & docPorServicioData(i, 1) & "-" & docPorServicioData(i, 2) & "-" & docPorServicioData(i, 3)) = docPorServicioData(i, 4)
    Next i

    diasMargen = wsDashboard.Range("G2").Value

    Dim updateCounter As Long
    updateCounter = 0

    For Each subcarpeta In carpetaProyecto.SubFolders
        placaVehiculo = subcarpeta.Name
        
        ' Filtrar por PLACA si se ha especificado
        If placa <> "" And placa <> placaVehiculo Then
            GoTo NextPLACA
        End If

        For Each subSubCarpeta In subcarpeta.SubFolders
            ' Crear una nueva fila en la tabla DashboardVehiculo
            Set fila = wsDashboard.ListObjects("DashboardEmbarcacion").ListRows.Add
            fila.Range(1, 1).Value = Proyecto

            ' Buscar información del vehículo
            If vehiculoDict.Exists(placaVehiculo) Then
                Dim vehiculoInfo As Variant
                vehiculoInfo = vehiculoDict(placaVehiculo)

                ' Filtrar por Empresa y Servicio si se han especificado
                 If (Empresa <> "" And Empresa <> vehiculoInfo(2)) Or (Servicio <> "" And Servicio <> vehiculoInfo(0)) Or (vehiculoInfo(0) <> "EMBARCACIÓN") Then
                    fila.Delete
                    GoTo NextSubSubCarpeta
                End If

                fila.Range(1, 2).Value = vehiculoInfo(2) ' Empresa
                fila.Range(1, 5).Value = vehiculoInfo(1) ' Servicio
                fila.Range(1, 4).Value = placaVehiculo ' PLACA
                fila.Range(1, 3).Value = vehiculoInfo(0) ' TIPO
                fila.Range(1, 12).Value = "Sin documento"
            End If

            ' Buscar información del documento
            Dim nombreCarpeta As String
            nombreCarpeta = subSubCarpeta.Name
            Dim keyDoc As String
            keyDoc = nombreCarpeta & "-" & fila.Range(1, 3).Value & "-" & fila.Range(1, 5).Value
            If docXServicioDict.Exists(keyDoc) Then
                fila.Range(1, 6).Value = docXServicioDict(keyDoc) ' Documento
                ' Link
                fila.Range(1, 11).Value = "Link"
                fila.Range(1, 11).Hyperlinks.Add Anchor:=fila.Range(1, 11), Address:=rutaRecurso & "\" & placaVehiculo & "\" & nombreCarpeta, TextToDisplay:="Link"
            End If

            ' Buscar información de la obligatoriedad del documento
            Dim keyObligatorio As String
            keyObligatorio = Proyecto & "-" & fila.Range(1, 3).Value & "-" & fila.Range(1, 5).Value & "-" & fila.Range(1, 6).Value
            If docPorServicioDict.Exists(keyObligatorio) Then
                fila.Range(1, 7).Value = docPorServicioDict(keyObligatorio) ' Obligatorio
            End If

            ' Buscar fechas de emisión y vigencia iterando
            Dim keyVehiculoDocs As String
            keyVehiculoDocs = placaVehiculo & "-" & fila.Range(1, 6).Value
            If vehiculoDocsDict.Exists(keyVehiculoDocs) Then
                Dim docsInfo As Variant
                docsInfo = vehiculoDocsDict(keyVehiculoDocs)
                Dim emision As Variant
                Dim vigencia As Variant
                emision = docsInfo(0)
                vigencia = docsInfo(1)
            
                
                fila.Range(1, 8).Value = emision ' Emisión
                fila.Range(1, 9).Value = vigencia ' Vigencia

                ' Insertar la fórmula en la columna de vigencia en días
                fila.Range(1, 10).Formula = "=IF([@VIGENCIA] >= TODAY(), [@VIGENCIA]-TODAY(), 0)"

                ' Calcular vigencia en días
                Dim vigenciaDias As Long
                vigenciaDias = DateDiff("d", Date, vigencia)

                ' Estado basado en vigencia
                If vigenciaDias >= diasMargen Then
                    fila.Range(1, 12).Value = "Vigente"
                ElseIf vigenciaDias >= 1 And vigenciaDias < diasMargen Then
                    fila.Range(1, 12).Value = "Por Vencer"
                Else
                    fila.Range(1, 12).Value = "Vencido"
                End If

                ' Observación
                fila.Range(1, 13).Value = docsInfo(2) ' Observación
                
                With fila.Range(1, 13)
                    ' Elimina cualquier validación previa en la celda
                    .Validation.Delete
                    ' Agrega la lista desplegable sin restricciones ni mensajes
                    .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                                    Operator:=xlBetween, Formula1:="APROBADO POR GERENCIA;NO RESTRICTIVO;NO APLICA;OBSERVADO"
                    .Validation.IgnoreBlank = True
                    .Validation.InCellDropdown = True
                    .Validation.ShowError = False ' Desactiva los mensajes de error para permitir cualquier texto
                    ' Establece el valor predeterminado de la celda
                    .Value = docsInfo(2) ' Asigna el valor de docsInfo(2) como observación
                End With
                Set cell = fila.Range(1, 11)
                 ' Verificar si hay un hipervínculo en la celda
                If cell.Hyperlinks.Count > 0 Then
                    link = cell.Hyperlinks(1).Address ' Captura el link del hipervínculo
                    
                   If Dir(link, vbDirectory) <> "" Then
                        folderPath = link & "\*.*" ' Obtener todos los archivos de la carpeta
                        Filename = Dir(folderPath)
            
                        ' Verificar si hay archivos en la carpeta
                        If Filename = "" Then
                            fila.Range(1, 12).Value = "Vacío" ' Coloca "Vacío" en la celda siguiente si la carpeta está vacía
                        End If
                    End If
                End If
            Else
                fila.Range(1, 10).Formula = "=IF(AND(ISNUMBER([@VIGENCIA]), [@VIGENCIA] >= TODAY()), [@VIGENCIA]-TODAY(), 0)"

            End If
            
            ' Contador para la actualización periódica
            updateCounter = updateCounter + 1
            If updateCounter Mod 100 = 0 Then
                Application.ScreenUpdating = True
                Application.ScreenUpdating = False
            End If
            
NextSubSubCarpeta:
        Next subSubCarpeta
NextPLACA:
    Next subcarpeta

    ' Restaurar el cálculo automático y la actualización de pantalla
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' Notificar al usuario
    If TipoRep <> 2 Then
        MsgBox "Dashboard Vehiculo generado correctamente", vbInformation
    End If
End Sub

Sub CopiarDatosAFilaActivaEmbarcacion()
    Dim wsDashboard As Worksheet
    Dim wsVehiculoDocs As Worksheet
    Dim filaActiva As ListRow
    Dim rngActiva As Range
    Dim codigoProyecto As String
    Dim DocIdentidad As String
    Dim documento As String
    Dim emisionStr As String
    Dim vigenciaStr As String
    Dim emision As Date
    Dim vigencia As Date
    Dim FechaDoc As Date
    Dim observacion As String
    Dim link As String
    Dim destinoArchivo As String
    Dim Abreviatura As String
    Dim filaDestino As Range
    
    Set wsDashboard = ThisWorkbook.Sheets("DASHBOARD EMBARCACIONES")
    Set wsVehiculoDocs = ThisWorkbook.Sheets("VEHICULO DOCUMENTOS")
    Set wsDocumentoMaster = ThisWorkbook.Sheets("MASTER REQUISITOS VEHICULO")
    
    ' Verificar si la celda activa está en la tabla DashboardPersonal
    On Error Resume Next
    Set rngActiva = Intersect(ActiveCell, wsDashboard.ListObjects("DashboardEmbarcacion").DataBodyRange)
    On Error GoTo 0
    
    If rngActiva Is Nothing Then
        MsgBox "Seleccione una celda dentro de la tabla DashboardEmbarcacion.", vbExclamation
        Exit Sub
    End If
    
    ' Obtener la fila activa de la tabla DashboardPersonal
    Set filaActiva = wsDashboard.ListObjects("DashboardEmbarcacion").ListRows(rngActiva.Row - wsDashboard.ListObjects("DashboardEmbarcacion").HeaderRowRange.Row)
    
    documento = filaActiva.Range(1, 6).Value
    Recurso = filaActiva.Range(1, 5).Value
    FechaDoc = filaActiva.Range(1, 8).Value

    ' Obtener la tabla Adjuntables de la hoja MASTER DOCUMENTOS PERSONAL
    Set tablaDocMaster = wsDocumentoMaster.ListObjects("RequisitosPorVehículo")

    ' Buscar el documentoAutorizado en la tabla Adjuntables (suponiendo que está en la columna 1)
    Set celdaEncontrada = tablaDocMaster.ListColumns(3).DataBodyRange.Find(documento, LookAt:=xlWhole)

    ' Si no se encuentra el documento en la tabla
    If celdaEncontrada Is Nothing Then
        MsgBox "El documento no está en la tabla Maestra", vbExclamation
       Exit Sub
    End If

    Abreviatura = celdaEncontrada.Offset(0, 4).Value
    
    emisionStr = filaActiva.Range(1, 8).Value
    vigenciaStr = filaActiva.Range(1, 9).Value
    
    If Not IsDate(emisionStr) Then
        MsgBox "La fecha de emisión no es un dato fecha, revisar", vbExclamation
        Exit Sub
    End If

    If Not IsDate(vigenciaStr) Then
        MsgBox "La fecha de vigencia no es un dato fecha, revisar", vbExclamation
        Exit Sub
    End If
    
    ' Convertir las cadenas a fechas
    emision = CDate(emisionStr)
    vigencia = CDate(vigenciaStr)

    ' Obtener los datos de la fila activa
    DocIdentidad = CStr(filaActiva.Range(1, 4).Value)
    documento = filaActiva.Range(1, 6).Value
    observacion = filaActiva.Range(1, 13).Value
    link = filaActiva.Range(1, 11).Hyperlinks(1).Address
    
    ' Verificar si EMISIÓN y VIGENCIA están presentes
    If IsEmpty(emision) Or IsEmpty(vigencia) Or emision = "00:00:00" Or vigencia = "00:00:00" Then
        MsgBox "Por favor ingrese los datos de EMISIÓN y VIGENCIA.", vbExclamation
        Exit Sub
    End If
   
       ' Buscar si ya existe la combinación de DOC IDENTIDAD y DOCUMENTO
    Dim encontrado As Range
    Set encontrado = wsVehiculoDocs.Columns(1).Find(What:=DocIdentidad, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not encontrado Is Nothing Then
        firstAddress = encontrado.Address
        Do
            If CStr(encontrado.Offset(0, 1).Value) = documento Then
                Set filaDestino = encontrado.EntireRow
                Exit Do
            End If
            Set encontrado = wsVehiculoDocs.Columns(1).FindNext(encontrado)
        Loop While Not encontrado Is Nothing And encontrado.Address <> firstAddress
    End If
        
    ' Si no se encontró, agregar una nueva fila
    If filaDestino Is Nothing Then
        Set filaDestino = wsVehiculoDocs.Rows(wsVehiculoDocs.Cells(wsVehiculoDocs.Rows.Count, "A").End(xlUp).Row + 1)
    End If
 
    ' Copiar los datos a la tabla PersonalDocumentos
    filaDestino.Cells(1, 1).Value = DocIdentidad
    filaDestino.Cells(1, 2).Value = documento
    filaDestino.Cells(1, 3).Value = emision
    filaDestino.Cells(1, 4).Value = vigencia
    filaDestino.Cells(1, 5).Value = observacion
    
        ' ---------------------------------------------------------------
    ' Ahora procesamos el libro origen (EXTERNO) y actualizamos los datos antes de mover el archivo
    ' ---------------------------------------------------------------

    ' Obtener la ruta del libro externo desde la celda B8 de la hoja PRINCIPAL
    rutaOrigen = ThisWorkbook.Sheets("PRINCIPAL").Range("B8").Value

    ' Verificar si la ruta del libro origen es válida
    If Dir(rutaOrigen) = "" Then
        MsgBox "El archivo de origen no se encuentra en la ruta especificada.", vbCritical
        Exit Sub
    End If

    ' Abrir el libro origen con permisos de escritura
    Set libroOrigen = Workbooks.Open(rutaOrigen, ReadOnly:=False) ' Permitimos edición
    
    ' Definir la hoja VEHICULO DOCUMENTOS del libro origen
    Set wsVehiculoDocsOrigen = libroOrigen.Sheets("VEHICULO DOCUMENTOS")

    ' Buscar si ya existe la combinación de CÓDIGO PROYECTO y DOC IDENTIDAD en el libro origen
    Set encontrado = wsVehiculoDocsOrigen.Columns(1).Find(What:=DocIdentidad, LookIn:=xlValues, LookAt:=xlWhole)

    If Not encontrado Is Nothing Then
        firstAddress = encontrado.Address
        Do
            If encontrado.Offset(0, 1).Value = documento Then
                Set filaDestinoOrigen = encontrado.EntireRow
                Exit Do
            End If
            Set encontrado = wsVehiculoDocsOrigen.Columns(1).FindNext(encontrado)
        Loop While Not encontrado Is Nothing And encontrado.Address <> firstAddress
    End If

    ' Si no se encontró en el libro origen, agregar una nueva fila
    If filaDestinoOrigen Is Nothing Then
        Set filaDestinoOrigen = wsVehiculoDocsOrigen.Rows(wsVehiculoDocsOrigen.Cells(wsVehiculoDocsOrigen.Rows.Count, "A").End(xlUp).Row + 1)
    End If

    ' Copiar los datos en la hoja VEHICULO DOCUMENTOS del libro origen
    filaDestinoOrigen.Cells(1, 1).Value = DocIdentidad
    filaDestinoOrigen.Cells(1, 2).Value = documento
    filaDestinoOrigen.Cells(1, 3).Value = emision
    filaDestinoOrigen.Cells(1, 4).Value = vigencia
    filaDestinoOrigen.Cells(1, 5).Value = observacion

    ' Cerrar el libro origen y guardar los cambios
    libroOrigen.Close SaveChanges:=True
    
    
    ' Mostrar la ventana de diálogo para seleccionar el archivo
    Dim archivoRuta As String
    archivoRuta = Application.GetOpenFilename("Archivos PDF o JPG (*.pdf;*.jpg), *.pdf;*.jpg", , "Seleccione el archivo a copiar")
    
    If archivoRuta <> "False" Then
        ' Verificar que la ruta en la columna LINK sea válida
        If Dir(link, vbDirectory) = "" Then
            MsgBox "La ruta especificada en la columna LINK no es válida: " & link, vbExclamation
            Exit Sub
        End If
        
        ' Construir la ruta completa del archivo de destino
        destinoArchivo = link & "\" & Dir(archivoRuta)
        
        Dim resultado As Boolean
        Dim AbrRecurso As String
        AbrRecurso = DocIdentidad
        resultado = CopiarArchivoConNuevoNombre(archivoRuta, _
                                                link, _
                                                Abreviatura, _
                                                AbrRecurso, _
                                                FechaDoc)
    
        ' Verificar si la copia fue exitosa
        If resultado Then
            MsgBox "Datos copiados y archivo transferido correctamente.", vbInformation
        Else
            MsgBox "Error al copiar el archivo."
        End If
        
        filaActiva.Range(1, 12).Value = EvaluarVigencia(vigencia)
    Else
        MsgBox "Datos copiados, pero no se seleccionó ningún archivo.", vbInformation
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al copiar el archivo: " & Err.Description, vbExclamation
End Sub

Function EvaluarVigenciaEmbarcacion(vigencia As Date) As String
    Dim wsDashboard As Worksheet
    Dim diasMargen As Integer
    Dim diasDiferencia As Integer
    
    Set wsDashboard = ThisWorkbook.Sheets("DASHBOARD EMBARCACIONES")
    
    ' Leer el valor de la celda G2
    diasMargen = wsDashboard.Range("G2").Value
    
    ' Calcular la diferencia entre la vigencia y la fecha actual
    diasDiferencia = DateDiff("d", Date, vigencia)
    
    ' Evaluar las condiciones
    If diasDiferencia >= diasMargen Then
        EvaluarVigenciaEmbarcacion = "Vigente"
    ElseIf diasDiferencia > 1 And diasDiferencia < diasMargen Then
        EvaluarVigenciaEmbarcacion = "Por Vencer"
    Else
        EvaluarVigenciaEmbarcacion = "Vencido"
    End If
End Function


Sub ActualizarVigenciaDiasEmbarcacion(rngActiva As Range)
    Dim wsDashboard As Worksheet
    Dim filaActiva As ListRow
    Dim vigencia As Date
    Dim vigenciaDias As Long
    
    Set wsDashboard = ThisWorkbook.Sheets("DASHBOARD EMBARCACIONES")
    
    ' Verificar si la celda activa está en la tabla DashboardEmbarcacion
    If rngActiva Is Nothing Then
        MsgBox "Seleccione una celda dentro de la tabla DashboardEmbarcacion.", vbExclamation
        Exit Sub
    End If
    
    ' Verificar si la fila activa es mayor o igual a la fila 11
    If rngActiva.Row < 11 Then
        Exit Sub
    End If
    
    ' Obtener la fila activa de la tabla DashboardPersonal
    Set filaActiva = wsDashboard.ListObjects("DashboardEmbarcacion").ListRows(rngActiva.Row - wsDashboard.ListObjects("DashboardEmbarcacion").HeaderRowRange.Row)
    
    ' Obtener la fecha de vigencia
    vigencia = filaActiva.Range(1, 9).Value ' Suponiendo que la columna VIGENCIA es la novena columna
    
    ' Calcular vigencia en días
    If Not IsEmpty(vigencia) Then
        vigenciaDias = DateDiff("d", Date, vigencia)
        filaActiva.Range(1, 10).Value = vigenciaDias ' Suponiendo que la columna VIGENCIA DIAS es la décima columna
    Else
        MsgBox "La columna VIGENCIA está vacía. Por favor ingrese una fecha válida.", vbExclamation
    End If
End Sub




