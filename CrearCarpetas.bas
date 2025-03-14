Attribute VB_Name = "CrearCarpetas"
Sub CrearEstructuraDeCarpetas()
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
    
    ' Registrar hora de inicio del proceso
    Dim horaInicio As Date
    horaInicio = Now()
    ThisWorkbook.Sheets("CATALOGOS").Range("N2").Value = Format(horaInicio, "dd/mm/yyyy hh:mm:ss")
    
    
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
       ' Registrar hora de finalización del proceso
    ThisWorkbook.Sheets("CATALOGOS").Range("N3").Value = Format(Now(), "dd/mm/yyyy hh:mm:ss")
    
    ThisWorkbook.Save
    MsgBox "Estructura de carpetas creada correctamente."
End Sub

Sub CrearEstructuraDeCarpetasMejorado()
    ' Desactivar actualizaciones de pantalla para acelerar
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

   ' Registrar hora de inicio del proceso
    Dim horaInicio As Date
    horaInicio = Now()
    ThisWorkbook.Sheets("CATALOGOS").Range("N2").Value = Format(horaInicio, "dd/mm/yyyy hh:mm:ss")
    
    Dim wsPrincipal As Worksheet, wsPersonal As Worksheet, wsDocumentosPuestos As Worksheet
    Dim wsVehiculos As Worksheet, wsRequisitosVehiculos As Worksheet
    Dim wsMasterDocumentos As Worksheet, wsMasterRequisitos As Worksheet
    Dim proyectoCodigo As String, carpetaRaiz As String
    Dim carpetaProyecto As String, carpetaHabilitaciones As String, carpetaRecursos As String
    Dim carpetaPersonal As String, carpetaVehiculos As String
    Dim personalData As Variant, vehiculosData As Variant, documentosPuestosData As Variant
    Dim masterDocumentosData As Variant, requisitosVehiculosData As Variant, masterRequisitosData As Variant
    Dim i As Long, j As Long, k As Long
    Dim documentoNombre As String, vehiculoNombre As String
    Dim linkPath As String, targetPath As String
    
    ' Diccionarios para búsquedas más rápidas
    Dim dictPuestosDocs As Object, dictDocsMaster As Object
    Dim dictReqVehiculos As Object, dictReqMaster As Object
    
    Set dictPuestosDocs = CreateObject("Scripting.Dictionary")
    Set dictDocsMaster = CreateObject("Scripting.Dictionary")
    Set dictReqVehiculos = CreateObject("Scripting.Dictionary")
    Set dictReqMaster = CreateObject("Scripting.Dictionary")
    
    ' Crear el objeto WScript.Shell (crearlo una sola vez)
    Dim ws As Object
    Set ws = CreateObject("WScript.Shell")
    
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
    Dim carpetasIniciales As Variant
    carpetasIniciales = Array( _
        carpetaProyecto, _
        carpetaHabilitaciones, _
        carpetaHabilitaciones & "\PERSONAL", _
        carpetaHabilitaciones & "\VEHICULOS", _
        carpetaRecursos, _
        carpetaRecursos & "\PERSONAL", _
        carpetaRecursos & "\VEHICULOS")
    
    Dim carpeta As Variant
    For Each carpeta In carpetasIniciales
        CrearCarpeta CStr(carpeta)
    Next carpeta
    
    ' Cargar los datos en arreglos para acceso más rápido
    personalData = wsPersonal.Range("A2:E" & wsPersonal.Cells(wsPersonal.Rows.Count, "A").End(xlUp).Row).Value
    documentosPuestosData = wsDocumentosPuestos.UsedRange.Value
    masterDocumentosData = wsMasterDocumentos.UsedRange.Value
    vehiculosData = wsVehiculos.Range("A2:E" & wsVehiculos.Cells(wsVehiculos.Rows.Count, "A").End(xlUp).Row).Value
    requisitosVehiculosData = wsRequisitosVehiculos.UsedRange.Value
    masterRequisitosData = wsMasterRequisitos.UsedRange.Value
    
    ' Preprocesar datos en diccionarios para búsqueda más rápida
    ' Crear diccionario para DOCUMENTO X PUESTO
    For j = 2 To UBound(documentosPuestosData, 1)
        Dim puestoKey As String
        puestoKey = CStr(documentosPuestosData(j, 1))
        
        If Not dictPuestosDocs.exists(puestoKey) Then
            Dim docsCollection As New Collection
            dictPuestosDocs.Add puestoKey, docsCollection
        End If
        
        dictPuestosDocs(puestoKey).Add CStr(documentosPuestosData(j, 2))
    Next j
    
    ' Crear diccionario para MASTER DOCUMENTOS PERSONAL
    For k = 2 To UBound(masterDocumentosData, 1)
        Dim docMasterKey As String
        docMasterKey = CStr(masterDocumentosData(k, 1)) & "|" & CStr(masterDocumentosData(k, 2))
        
        If Not dictDocsMaster.exists(docMasterKey) Then
            dictDocsMaster.Add docMasterKey, CStr(masterDocumentosData(k, 3))
        End If
    Next k
    
    ' Crear diccionario para REQUISITOS X VEHICULO
    For j = 2 To UBound(requisitosVehiculosData, 1)
        If CStr(requisitosVehiculosData(j, 5)) = proyectoCodigo Then
            Dim reqVehKey As String
            reqVehKey = CStr(requisitosVehiculosData(j, 1)) & "|" & CStr(requisitosVehiculosData(j, 2))
            
            If Not dictReqVehiculos.exists(reqVehKey) Then
                Dim reqCollection As New Collection
                dictReqVehiculos.Add reqVehKey, reqCollection
            End If
            
            dictReqVehiculos(reqVehKey).Add CStr(requisitosVehiculosData(j, 3))
        End If
    Next j
    
    ' Crear diccionario para MASTER REQUISITOS VEHICULO
    For k = 2 To UBound(masterRequisitosData, 1)
        Dim reqMasterKey As String
        reqMasterKey = CStr(masterRequisitosData(k, 1)) & "|" & CStr(masterRequisitosData(k, 2)) & "|" & CStr(masterRequisitosData(k, 3))
        
        If Not dictReqMaster.exists(reqMasterKey) Then
            dictReqMaster.Add reqMasterKey, CStr(masterRequisitosData(k, 4))
        End If
    Next k
    
    ' Crear carpetas para Personal (utilizando los diccionarios)
    For i = 1 To UBound(personalData, 1)
        If CStr(personalData(i, 5)) = proyectoCodigo Then
            documentoNombre = Trim(CStr(personalData(i, 1)))
            carpetaPersonal = carpetaHabilitaciones & "\PERSONAL\" & documentoNombre
            CrearCarpeta carpetaPersonal
            
            Dim puesto As String
            puesto = CStr(personalData(i, 3))
            
            ' Verificar si el puesto tiene documentos asociados
            If dictPuestosDocs.exists(puesto) Then
                Dim documento As Variant
                
                ' Iterar sobre los documentos asociados al puesto
                For Each documento In dictPuestosDocs(puesto)
                    Dim docKey As String
                    docKey = puesto & "|" & documento
                    
                    ' Verificar si el documento está en el master
                    If dictDocsMaster.exists(docKey) Then
                        Dim nombreDoc As String
                        nombreDoc = dictDocsMaster(docKey)
                        
                        ' Crear las carpetas necesarias
                        CrearCarpeta carpetaPersonal & "\" & nombreDoc
                        CrearCarpeta carpetaRecursos & "\PERSONAL\" & documentoNombre
                        CrearCarpeta carpetaRecursos & "\PERSONAL\" & documentoNombre & "\" & nombreDoc
                        
                        ' Crear el acceso directo
                        linkPath = carpetaPersonal & "\" & nombreDoc & ".lnk"
                        targetPath = carpetaRecursos & "\PERSONAL\" & documentoNombre & "\" & nombreDoc
                        
                        With ws.CreateShortcut(linkPath)
                            .targetPath = targetPath
                            .Save
                        End With
                    End If
                Next documento
            End If
        End If
    Next i
    
    ' Crear carpetas para Vehículos (utilizando los diccionarios)
    For i = 1 To UBound(vehiculosData, 1)
        If CStr(vehiculosData(i, 5)) = proyectoCodigo Then
            vehiculoNombre = CStr(vehiculosData(i, 1))
            carpetaVehiculos = carpetaHabilitaciones & "\VEHICULOS\" & vehiculoNombre
            CrearCarpeta carpetaVehiculos
            
            Dim tipoVeh As String, categoriaVeh As String
            tipoVeh = CStr(vehiculosData(i, 2))
            categoriaVeh = CStr(vehiculosData(i, 3))
            
            Dim vehKey As String
            vehKey = tipoVeh & "|" & categoriaVeh
            
            ' Verificar si el tipo y categoría de vehículo tienen requisitos asociados
            If dictReqVehiculos.exists(vehKey) Then
                Dim requisito As Variant
                
                ' Iterar sobre los requisitos asociados al vehículo
                For Each requisito In dictReqVehiculos(vehKey)
                    Dim reqKey As String
                    reqKey = tipoVeh & "|" & categoriaVeh & "|" & requisito
                    
                    ' Verificar si el requisito está en el master
                    If dictReqMaster.exists(reqKey) Then
                        Dim nombreReq As String
                        nombreReq = dictReqMaster(reqKey)
                        
                        ' Crear las carpetas necesarias
                        CrearCarpeta carpetaVehiculos & "\" & nombreReq
                        CrearCarpeta carpetaRecursos & "\VEHICULOS\" & vehiculoNombre
                        CrearCarpeta carpetaRecursos & "\VEHICULOS\" & vehiculoNombre & "\" & nombreReq
                        
                        ' Crear el acceso directo
                        linkPath = carpetaVehiculos & "\" & nombreReq & ".lnk"
                        targetPath = carpetaRecursos & "\VEHICULOS\" & vehiculoNombre & "\" & nombreReq
                        
                        With ws.CreateShortcut(linkPath)
                            .targetPath = targetPath
                            .Save
                        End With
                    End If
                Next requisito
            End If
        End If
    Next i
    
    ' Restaurar la configuración de Excel
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    ThisWorkbook.Sheets("CATALOGOS").Range("N3").Value = Format(Now(), "dd/mm/yyyy hh:mm:ss")
    
    ' Guardar el libro antes de mostrar el mensaje
    ThisWorkbook.Save
    MsgBox "Estructura de carpetas creada correctamente."
End Sub

Sub CrearCarpeta(ByVal ruta As String)
    If Dir(ruta, vbDirectory) = "" Then
        MkDir ruta
    End If
End Sub

