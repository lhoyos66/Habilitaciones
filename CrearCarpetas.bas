Attribute VB_Name = "CrearCarpetas"
Sub CrearEstructuraDeCarpetasSUPERADO()
    Dim wsPrincipal As Worksheet, wsPersonal As Worksheet, wsDocumentosPuestos As Worksheet
    Dim wsVehiculos As Worksheet, wsRequisitosVehiculos As Worksheet
    Dim wsMasterDocumentos As Worksheet, wsMasterRequisitos As Worksheet
    Dim proyectoCodigo As String, carpetaRaiz As String
    Dim carpetaProyecto As String, carpetaHabilitaciones As String, carpetaRecursos As String
    Dim carpetaPersonal As String, carpetaVehiculos As String
    Dim personalRango As Range, vehiculosRango As Range
    Dim personalData As Variant, vehiculosData As Variant, documentosPuestosData As Variant
    Dim masterDocumentosData As Variant, requisitosVehiculosData As Variant, masterRequisitosData As Variant
    Dim i As Long, j As Long, k As Long
    Dim documentoNombre As String, vehiculoNombre As String
    Dim ruta As String
    ' Crear el objeto WScript.Shell para crear el acceso directo
    Dim ws As Object
    Set ws = CreateObject("WScript.Shell")
    ' Crear un link en carpetaPersonal
    Dim linkPath As String
    Dim targetPath As String
    ' Crear el acceso directo
    Dim shortcut As Object
    
    
    ' Establecer las hojas
    Set wsPrincipal = ThisWorkbook.Sheets("PRINCIPAL")
    Set wsPersonal = ThisWorkbook.Sheets("PERSONAL")
    Set wsDocumentosPuestos = ThisWorkbook.Sheets("DOCUMENTO X PUESTO")
    Set wsVehiculos = ThisWorkbook.Sheets("VEHICULOS")
    Set wsRequisitosVehiculos = ThisWorkbook.Sheets("REQUISITOS X VEHICULO")
    Set wsMasterDocumentos = ThisWorkbook.Sheets("MASTER DOCUMENTOS PERSONAL")
    Set wsMasterRequisitos = ThisWorkbook.Sheets("MASTER REQUISITOS VEHICULO")
    
    ' Obtener valores de la hoja PRINCIPAL
    proyectoCodigo = CStr(wsPrincipal.Range("A2").Value)
    carpetaRaiz = wsPrincipal.Range("B6").Value
    carpetaProyecto = carpetaRaiz & "\" & proyectoCodigo
    carpetaHabilitaciones = carpetaProyecto & "\HABILITACIONES"
    carpetaRecursos = carpetaRaiz & "\RECURSOS"
    
    ' Crear carpetas iniciales
    CrearCarpeta carpetaProyecto
    CrearCarpeta carpetaHabilitaciones
    CrearCarpeta carpetaHabilitaciones & "\PERSONAL"
    CrearCarpeta carpetaHabilitaciones & "\VEHICULOS"
    CrearCarpeta carpetaRecursos
    CrearCarpeta carpetaRecursos & "\PERSONAL"
    CrearCarpeta carpetaRecursos & "\VEHICULOS"
    
    ' Cargar los datos en arreglos para acceso más rápido
    personalData = wsPersonal.Range("A2:E" & wsPersonal.Cells(wsPersonal.Rows.Count, "A").End(xlUp).Row).Value
    documentosPuestosData = wsDocumentosPuestos.UsedRange.Value
    masterDocumentosData = wsMasterDocumentos.UsedRange.Value
    vehiculosData = wsVehiculos.Range("A2:E" & wsVehiculos.Cells(wsVehiculos.Rows.Count, "A").End(xlUp).Row).Value
    requisitosVehiculosData = wsRequisitosVehiculos.UsedRange.Value
    masterRequisitosData = wsMasterRequisitos.UsedRange.Value
    
    ' Crear carpetas para Personal
    For i = 1 To UBound(personalData, 1)
        If CStr(personalData(i, 5)) = proyectoCodigo Then
            documentoNombre = Trim(personalData(i, 1))
            carpetaPersonal = carpetaHabilitaciones & "\PERSONAL\" & documentoNombre
            CrearCarpeta carpetaPersonal
            
            ' Buscar documentos relacionados al puesto
            For j = 2 To UBound(documentosPuestosData, 1)
                If documentosPuestosData(j, 1) = personalData(i, 3) Then
                    For k = 2 To UBound(masterDocumentosData, 1)
                        If masterDocumentosData(k, 1) = personalData(i, 3) And _
                           masterDocumentosData(k, 2) = documentosPuestosData(j, 2) Then
            
                            CrearCarpeta carpetaPersonal & "\" & masterDocumentosData(k, 3)
                            CrearCarpeta carpetaRecursos & "\PERSONAL\" & documentoNombre
                            CrearCarpeta carpetaRecursos & "\PERSONAL\" & documentoNombre & "\" & masterDocumentosData(k, 3)
                            
                            linkPath = carpetaPersonal & "\" & masterDocumentosData(k, 3) & ".lnk"
                            targetPath = carpetaRecursos & "\PERSONAL\" & documentoNombre & "\" & masterDocumentosData(k, 3)
                            
                            Set shortcut = ws.CreateShortcut(linkPath)
                            shortcut.targetPath = targetPath
                            shortcut.Save
            
                        End If
                    Next k
                End If
            Next j

        End If
    Next i
    
    ' Crear carpetas para Vehículos
    For i = 1 To UBound(vehiculosData, 1)
        If CStr(vehiculosData(i, 5)) = proyectoCodigo Then
            vehiculoNombre = vehiculosData(i, 1)
            carpetaVehiculos = carpetaHabilitaciones & "\VEHICULOS\" & vehiculoNombre
            CrearCarpeta carpetaVehiculos
            
            ' Buscar requisitos relacionados al vehículo
            For j = 2 To UBound(requisitosVehiculosData, 1)
                If requisitosVehiculosData(j, 1) = vehiculosData(i, 2) And _
                   requisitosVehiculosData(j, 2) = vehiculosData(i, 3) And _
                   requisitosVehiculosData(j, 5) = proyectoCodigo Then
                    For k = 2 To UBound(masterRequisitosData, 1)
                        If masterRequisitosData(k, 1) = vehiculosData(i, 2) And _
                           masterRequisitosData(k, 2) = vehiculosData(i, 3) And _
                           masterRequisitosData(k, 3) = requisitosVehiculosData(j, 3) Then
                            CrearCarpeta carpetaVehiculos & "\" & masterRequisitosData(k, 4)
                            CrearCarpeta carpetaRecursos & "\VEHICULOS\" & vehiculoNombre
                            CrearCarpeta carpetaRecursos & "\VEHICULOS\" & vehiculoNombre & "\" & masterRequisitosData(k, 4)
                            
                            linkPath = carpetaVehiculos & "\" & masterRequisitosData(k, 4) & ".lnk"
                            targetPath = carpetaRecursos & "\VEHICULOS\" & vehiculoNombre & "\" & masterRequisitosData(k, 4)
                            
                            Set shortcut = ws.CreateShortcut(linkPath)
                            shortcut.targetPath = targetPath
                            shortcut.Save
                        End If
                    Next k
                End If
            Next j
        End If
    Next i
    'EliminarCarpetasInnecesarias (1)
    ' Guardar el libro antes de mostrar el mensaje
    ThisWorkbook.Save
    MsgBox "Estructura de carpetas creada correctamente."
End Sub

Private Sub CrearCarpeta(ByVal ruta As String) ' Superado
    If Dir(ruta, vbDirectory) = "" Then
        MkDir ruta
    End If
End Sub

Sub CrearEstructuraDeCarpetas()
    Dim wsPrincipal As Worksheet
    Dim wsPersonal As Worksheet, wsDocumentosPuestos As Worksheet
    Dim wsVehiculos As Worksheet, wsRequisitosVehiculos As Worksheet
    Dim wsMasterDocumentos As Worksheet, wsMasterRequisitos As Worksheet
    Dim proyectoCodigo As String, carpetaRaiz As String
    Dim carpetaProyecto As String, carpetaHabilitaciones As String, carpetaRecursos As String
    Dim personalData As Variant, vehiculosData As Variant, documentosPuestosData As Variant
    Dim masterDocumentosData As Variant, requisitosVehiculosData As Variant, masterRequisitosData As Variant
    Dim i As Long
    
    ' Definir ruta del archivo de depuración
    rutaArchivoDepuracion = "C:\Users\lhoyos\Downloads\Depuracion_Carpetas.txt" ' Cambia la ruta si es necesario
    archivo = FreeFile

    ' Crear o sobrescribir el archivo de depuración
    Open rutaArchivoDepuracion For Output As archivo
    Print #archivo, "Inicio de Depuración - Creación de Carpetas"
    Print #archivo, "------------------------------------------------------------"

    ' Establecer las hojas de trabajo
    Set wsPrincipal = ThisWorkbook.Sheets("PRINCIPAL")
    Set wsPersonal = ThisWorkbook.Sheets("PERSONAL")
    Set wsDocumentosPuestos = ThisWorkbook.Sheets("DOCUMENTO X PUESTO")
    Set wsVehiculos = ThisWorkbook.Sheets("VEHICULOS")
    Set wsRequisitosVehiculos = ThisWorkbook.Sheets("REQUISITOS X VEHICULO")
    Set wsMasterDocumentos = ThisWorkbook.Sheets("MASTER DOCUMENTOS PERSONAL")
    Set wsMasterRequisitos = ThisWorkbook.Sheets("MASTER REQUISITOS VEHICULO")
    
    ' Obtener valores de la hoja PRINCIPAL
    proyectoCodigo = CStr(wsPrincipal.Range("A2").Value)
    carpetaRaiz = wsPrincipal.Range("B6").Value
    carpetaProyecto = carpetaRaiz & "\" & proyectoCodigo
    carpetaHabilitaciones = carpetaProyecto & "\HABILITACIONES"
    carpetaRecursos = carpetaRaiz & "\RECURSOS"
    
    ' Crear carpetas base
    CrearCarpetas Array(carpetaProyecto, carpetaHabilitaciones, carpetaHabilitaciones & "\PERSONAL", _
                        carpetaHabilitaciones & "\VEHICULOS", carpetaRecursos, carpetaRecursos & "\PERSONAL", _
                        carpetaRecursos & "\VEHICULOS")
    
    ' Depuración: Verificar que las carpetas base se han creado
    Print #archivo, "Carpetas base creadas:"
    Debug.Print "Carpetas base creadas:"
    Dim carpeta As Variant
    For Each carpeta In Array(carpetaProyecto, carpetaHabilitaciones, carpetaHabilitaciones & "\PERSONAL", _
                              carpetaHabilitaciones & "\VEHICULOS", carpetaRecursos, carpetaRecursos & "\PERSONAL", _
                              carpetaRecursos & "\VEHICULOS")
        Print #archivo, "- " & carpeta
        Debug.Print "- " & carpeta
    Next carpeta
        
    ' Cargar los datos en arreglos para mejorar rendimiento
    personalData = ObtenerDatos(wsPersonal, "A2:E")
    documentosPuestosData = wsDocumentosPuestos.UsedRange.Value
    masterDocumentosData = wsMasterDocumentos.UsedRange.Value
    vehiculosData = ObtenerDatos(wsVehiculos, "A2:E")
    requisitosVehiculosData = wsRequisitosVehiculos.UsedRange.Value
    masterRequisitosData = wsMasterRequisitos.UsedRange.Value
    
    ' Crear carpetas para Personal y Vehículos con accesos directos
    For i = 1 To UBound(personalData, 1)
        If CStr(personalData(i, 5)) = proyectoCodigo Then
            Print #archivo, "Creando estructura para personal: " & personalData(i, 1)
            Debug.Print "Creando estructura para personal: " & personalData(i, 1)
            CrearEstructuraPersonal carpetaHabilitaciones, carpetaRecursos, documentosPuestosData, masterDocumentosData, _
                                    personalData(i, 1), personalData(i, 3)
        End If
    Next i

    For i = 1 To UBound(vehiculosData, 1)
        If CStr(vehiculosData(i, 5)) = proyectoCodigo Then
            Print #archivo, "Creando estructura para vehículo: " & vehiculosData(i, 1)
            Debug.Print "Creando estructura para vehículo: " & vehiculosData(i, 1)
            CrearEstructuraVehiculos carpetaHabilitaciones, carpetaRecursos, requisitosVehiculosData, masterRequisitosData, _
                                     vehiculosData(i, 1), vehiculosData(i, 2), vehiculosData(i, 3)
        End If
    Next i
    
    ' Cerrar archivo de depuración
    Print #archivo, "------------------------------------------------------------"
    Print #archivo, "Fin de Depuración"
    Close archivo
    
    ' Guardar el libro y mostrar mensaje de confirmación
    ThisWorkbook.Save
    MsgBox "Estructura de carpetas creada correctamente, incluyendo accesos directos."
End Sub

' ?? Función para crear carpetas en un array
Sub CrearCarpetas(ByVal rutas As Variant)
    Dim ruta As Variant
    For Each ruta In rutas
        If Dir(ruta, vbDirectory) = "" Then MkDir ruta
    Next ruta
End Sub

' ?? Función para obtener datos en un rango dinámico
Function ObtenerDatos(ByVal ws As Worksheet, ByVal rangoBase As String) As Variant
    Dim rangoFinal As Range
    Set rangoFinal = ws.Range(rangoBase & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
    ObtenerDatos = rangoFinal.Value
End Function

' ?? Función para crear estructura de carpetas de Personal con accesos directos
Sub CrearEstructuraPersonal(ByVal carpetaHabilitaciones As String, ByVal carpetaRecursos As String, _
                            ByVal documentosPuestosData As Variant, ByVal masterDocumentosData As Variant, _
                            ByVal documentoNombre As String, ByVal puestoID As String)
    
    Dim carpetaPersonal As String, j As Long, k As Long
    carpetaPersonal = carpetaHabilitaciones & "\PERSONAL\" & documentoNombre
    CrearCarpetas Array(carpetaPersonal, carpetaRecursos & "\PERSONAL\" & documentoNombre)
    
    ' Crear carpetas y accesos directos para documentos de personal
    For j = 2 To UBound(documentosPuestosData, 1)
        If documentosPuestosData(j, 1) = puestoID Then
            For k = 2 To UBound(masterDocumentosData, 1)
                If masterDocumentosData(k, 1) = puestoID And masterDocumentosData(k, 2) = documentosPuestosData(j, 2) Then
                    Dim carpetaDestino As String
                    carpetaDestino = carpetaRecursos & "\PERSONAL\" & documentoNombre & "\" & masterDocumentosData(k, 3)
                    CrearCarpetas Array(carpetaPersonal & "\" & masterDocumentosData(k, 3), carpetaDestino)
                    
                    ' Crear acceso directo
                    CrearAccesoDirecto carpetaPersonal & "\" & masterDocumentosData(k, 3) & ".lnk", carpetaDestino
                End If
            Next k
        End If
    Next j
End Sub

' ?? Función para crear estructura de carpetas de Vehículos con accesos directos
Sub CrearEstructuraVehiculos(ByVal carpetaHabilitaciones As String, ByVal carpetaRecursos As String, _
                             ByVal requisitosVehiculosData As Variant, ByVal masterRequisitosData As Variant, _
                             ByVal vehiculoNombre As String, ByVal tipoVehiculo As String, ByVal categoria As String)
    
    Dim carpetaVehiculos As String, j As Long, k As Long
    carpetaVehiculos = carpetaHabilitaciones & "\VEHICULOS\" & vehiculoNombre
    CrearCarpetas Array(carpetaVehiculos, carpetaRecursos & "\VEHICULOS\" & vehiculoNombre)
    
    ' Crear carpetas y accesos directos para requisitos de vehículos
    For j = 2 To UBound(requisitosVehiculosData, 1)
        If requisitosVehiculosData(j, 1) = tipoVehiculo And requisitosVehiculosData(j, 2) = categoria Then
            For k = 2 To UBound(masterRequisitosData, 1)
                If masterRequisitosData(k, 1) = tipoVehiculo And masterRequisitosData(k, 2) = categoria And _
                   masterRequisitosData(k, 3) = requisitosVehiculosData(j, 3) Then
                    Dim carpetaDestino As String
                    carpetaDestino = carpetaRecursos & "\VEHICULOS\" & vehiculoNombre & "\" & masterRequisitosData(k, 4)
                    CrearCarpetas Array(carpetaVehiculos & "\" & masterRequisitosData(k, 4), carpetaDestino)
                    
                    ' Crear acceso directo
                    CrearAccesoDirecto carpetaVehiculos & "\" & masterRequisitosData(k, 4) & ".lnk", carpetaDestino
                End If
            Next k
        End If
    Next j
End Sub

' ?? Función para crear accesos directos
Sub CrearAccesoDirecto(ByVal linkPath As String, ByVal targetPath As String)
    Dim ws As Object, shortcut As Object
    Set ws = CreateObject("WScript.Shell")
    Set shortcut = ws.CreateShortcut(linkPath)
    shortcut.targetPath = targetPath
    shortcut.Save
End Sub


