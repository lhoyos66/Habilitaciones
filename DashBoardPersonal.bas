Attribute VB_Name = "DashBoardPersonal"
Sub EjecutarGenerarDashboardPersonal()
    Dim wsDashboard As Worksheet
    Dim Proyecto As String
    Dim Empresa As String
    Dim ServicioOPuesto As String
    Dim DNI As String
    
    Set wsDashboard = ThisWorkbook.Sheets("DASHBOARD PERSONAL")
    
    ' Obtener valores de las celdas específicas
    Proyecto = wsDashboard.Range("C2").Value
    Empresa = wsDashboard.Range("C4").Value
    ServicioOPuesto = wsDashboard.Range("C6").Value
    DNI = wsDashboard.Range("C8").Value
    
    ' Verificar si el proyecto ha sido ingresado
    If Proyecto = "" Then
        MsgBox "El proyecto es un campo obligatorio.", vbExclamation
        Exit Sub
    End If
    
    ' Llamar a la macro principal con los parámetros obtenidos
    Call GenerarDashboardPersonal(Proyecto, 1, Empresa, ServicioOPuesto, DNI)
End Sub

Sub EjecutarPersonalResumen()
    Dim wsDashboard As Worksheet
    Dim Proyecto As String
    Dim Empresa As String
    Dim ServicioOPuesto As String
    Dim DNI As String
    
    Set wsDashboard = ThisWorkbook.Sheets("PERSONAL RESUMEN")
    
    ' Obtener valores de las celdas específicas
    Proyecto = wsDashboard.Range("C2").Value
    Empresa = wsDashboard.Range("C4").Value
    ServicioOPuesto = wsDashboard.Range("C6").Value
    DNI = wsDashboard.Range("C8").Value
    
    ' Verificar si el proyecto ha sido ingresado
    If Proyecto = "" Then
        MsgBox "El proyecto es un campo obligatorio.", vbExclamation
        Exit Sub
    End If
    
    ' Llamar a la macro principal con los parámetros obtenidos
    Call GenerarResumenPersonal(Proyecto, Empresa, ServicioOPuesto, DNI)
End Sub

Sub DashBoardPersonalDesdeResumen()
    Dim wsPersonalResumen As Worksheet
    Dim wsDashboard As Worksheet
    Dim tblPersonalResumen As ListObject
    Dim filaActiva As ListRow
    Dim Proyecto As String
    Dim Empresa As String
    Dim ServicioOPuesto As String
    Dim DNI As String

    Set wsPersonalResumen = ThisWorkbook.Sheets("PERSONAL RESUMEN")
    Set wsDashboard = ThisWorkbook.Sheets("DASHBOARD PERSONAL")
    
    Set tblPersonalResumen = wsPersonalResumen.ListObjects("PersonalResumen")
 
    ' Verificar si la celda activa está dentro de la tabla PersonalResumen
    If Not Intersect(ActiveCell, tblPersonalResumen.DataBodyRange) Is Nothing Then
        ' Obtener la fila activa de la tabla
        Set filaActiva = tblPersonalResumen.ListRows(ActiveCell.Row - tblPersonalResumen.HeaderRowRange.Row)

        ' Capturar los valores de la fila activa
        Proyecto = filaActiva.Range(1, 1).Value
        Empresa = filaActiva.Range(1, 2).Value
        ServicioOPuesto = filaActiva.Range(1, 3).Value
        DNI = filaActiva.Range(1, 4).Value

        ' Copiar los valores a las celdas correspondientes en la hoja DASHBOARD PERSONAL
        wsDashboard.Range("C2").Value = Proyecto
        wsDashboard.Range("C4").Value = Empresa
        wsDashboard.Range("C6").Value = ServicioOPuesto
        wsDashboard.Range("C8").Value = "'" & DNI

        ' Verificar si la hoja DASHBOARD PERSONAL está oculta y mostrarla si es necesario
        If wsDashboard.Visible = xlSheetHidden Then
            wsDashboard.Visible = xlSheetVisible
        End If

        ' Activar la hoja DASHBOARD PERSONAL
        wsDashboard.Activate

        ' Ejecutar la macro GenerarDashboardPersonal
        Call GenerarDashboardPersonal(Proyecto, 2, Empresa, ServicioOPuesto, DNI)
    Else
        MsgBox "Por favor, seleccione una celda dentro de la tabla PersonalResumen.", vbExclamation
    End If
End Sub

Sub CopiarDatosAFilaActivaPersonal()
    Application.ScreenUpdating = False ' Desactivar la actualización de pantalla
    Application.DisplayAlerts = False ' Desactivar alertas
    Application.Calculation = xlCalculationManual ' Desactivar cálculos automáticos
    
    On Error GoTo ErrorHandler

    Dim wsDashboard As Worksheet
    Dim wsPersonalDocs As Worksheet
    Dim filaActiva As ListRow
    Dim rngActiva As Range
    Dim codigoProyecto As String
    Dim DocIdentidad As String
    Dim documento As String
    Dim emisionStr As String
    Dim vigenciaStr As String
    Dim emision As Date
    Dim vigencia As Date
    Dim observacion As String
    Dim link As String
    Dim destinoArchivo As String
    Dim tablaDocMaster As ListObject
    Dim celdaEncontrada As Range
    Dim Recurso As String
    Dim Abreviatura As String
    Dim FechaDoc As Date
    Dim encontrado As Range
    Dim filaDestino As Range
    Dim firstAddress As String
    Dim filaDestinoOrigen As Range
    
    Set wsDashboard = ThisWorkbook.Sheets("DASHBOARD PERSONAL")
    Set wsPersonalDocs = ThisWorkbook.Sheets("PERSONAL DOCUMENTOS")
    Set wsDocumentoMaster = ThisWorkbook.Sheets("MASTER DOCUMENTOS PERSONAL")
    Set wsCatalogos = ThisWorkbook.Sheets("CATALOGOS")

    ' Verificar si la celda activa está en la tabla DashboardPersonal
    On Error Resume Next
    Set rngActiva = Intersect(ActiveCell, wsDashboard.ListObjects("DashboardPersonal").DataBodyRange)
    On Error GoTo 0
    
    If rngActiva Is Nothing Then
        MsgBox "Seleccione una celda dentro de la tabla DashboardPersonal.", vbExclamation
        Exit Sub
    End If
    
    ' Obtener la fila activa de la tabla DashboardPersonal (LIBRO ACTIVO)
    Set filaActiva = wsDashboard.ListObjects("DashboardPersonal").ListRows(rngActiva.Row - wsDashboard.ListObjects("DashboardPersonal").HeaderRowRange.Row)

    documento = filaActiva.Range(1, 6).Value
    Recurso = filaActiva.Range(1, 5).Value
    FechaDoc = filaActiva.Range(1, 8).Value

    ' Obtener la tabla Adjuntables de la hoja MASTER DOCUMENTOS PERSONAL
    Set tablaDocMaster = wsDocumentoMaster.ListObjects("DocumentosXPuesto")

    ' Buscar el documentoAutorizado en la tabla Adjuntables (suponiendo que está en la columna 1)
    Set celdaEncontrada = tablaDocMaster.ListColumns(2).DataBodyRange.Find(documento, LookAt:=xlWhole)

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
    documento = CStr(filaActiva.Range(1, 6).Value)
    observacion = filaActiva.Range(1, 13).Value
    link = filaActiva.Range(1, 11).Hyperlinks(1).Address
    
    ' Verificar si EMISIÓN y VIGENCIA están presentes
    If IsEmpty(emisionStr) Or IsEmpty(vigenciaStr) Or emisionStr = "00:00:00" Or vigenciaStr = "00:00:00" Then
        MsgBox "Por favor ingrese los datos de EMISIÓN y VIGENCIA.", vbExclamation
        Exit Sub
    End If
    
    ' Buscar si ya existe la combinación de DOC IDENTIDAD y DOCUMENTO
    Set encontrado = wsPersonalDocs.Columns(1).Find(What:=DocIdentidad, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not encontrado Is Nothing Then
        firstAddress = encontrado.Address
        Do
            If CStr(encontrado.Offset(0, 1).Value) = documento Then
                Set filaDestino = encontrado.EntireRow
                Exit Do
            End If
            Set encontrado = wsPersonalDocs.Columns(1).FindNext(encontrado)
        Loop While Not encontrado Is Nothing And encontrado.Address <> firstAddress
    End If
        
    ' Si no se encontró, agregar una nueva fila
    If filaDestino Is Nothing Then
        Set filaDestino = wsPersonalDocs.Rows(wsPersonalDocs.Cells(wsPersonalDocs.Rows.Count, "A").End(xlUp).Row + 1)
    End If
    
    ' Copiar los datos a la tabla PersonalDocumentos
    filaDestino.Cells(1, 1).Value = DocIdentidad
    filaDestino.Cells(1, 2).Value = documento
    filaDestino.Cells(1, 3).Value = emision
    filaDestino.Cells(1, 4).Value = vigencia
    filaDestino.Cells(1, 5).Value = observacion
    
' ---------------------------------------------------------------
    ' Ahora procesamos el libro origen (EXTERNO)
    ' ---------------------------------------------------------------
    Application.EnableEvents = False ' Desactiva eventos para evitar disparadores innecesarios

    ' Obtener la ruta UNC del libro externo desde la celda B8 de la hoja PRINCIPAL
    rutaOrigen = ThisWorkbook.Sheets("PRINCIPAL").Range("B8").Value

    ' Verificar si la ruta del libro origen es válida
    If Dir(rutaOrigen) = "" Then
        MsgBox "El archivo de origen no se encuentra en la ruta especificada.", vbCritical
        Exit Sub
    End If

    ' Abrir el libro origen con permisos de escritura (ReadOnly=False)
    Set libroOrigen = Workbooks.Open(rutaOrigen, ReadOnly:=False) ' Permitimos edición
    
    ' Definir la hoja PERSONAL DOCUMENTOS del libro origen
    Set wsPersonalDocsOrigen = libroOrigen.Sheets("PERSONAL DOCUMENTOS")

    ' Buscar si ya existe la combinación de DOC IDENTIDAD y DOCUMENTO en el libro origen
    Set encontrado = wsPersonalDocsOrigen.Columns(1).Find(What:=DocIdentidad, LookIn:=xlValues, LookAt:=xlWhole)

    If Not encontrado Is Nothing Then
        firstAddress = encontrado.Address
        Do
            If CStr(encontrado.Offset(0, 1).Value) = documento Then
                Set filaDestinoOrigen = encontrado.EntireRow
                Exit Do
            End If
            Set encontrado = wsPersonalDocsOrigen.Columns(1).FindNext(encontrado)
        Loop While Not encontrado Is Nothing And encontrado.Address <> firstAddress
    End If

    ' Si no se encontró en el libro origen, agregar una nueva fila
    If filaDestinoOrigen Is Nothing Then
        Set filaDestinoOrigen = wsPersonalDocsOrigen.Rows(wsPersonalDocsOrigen.Cells(wsPersonalDocsOrigen.Rows.Count, "A").End(xlUp).Row + 1)
    End If

    ' Copiar los datos en la hoja PERSONAL DOCUMENTOS del libro origen
    filaDestinoOrigen.Cells(1, 1).Value = DocIdentidad
    filaDestinoOrigen.Cells(1, 2).Value = documento
    filaDestinoOrigen.Cells(1, 3).Value = emision
    filaDestinoOrigen.Cells(1, 4).Value = vigencia
    filaDestinoOrigen.Cells(1, 5).Value = observacion
    filaDestinoOrigen.Cells(1, 6).Value = 0

    ' Cerrar el libro origen en background y guardar los cambios sin notificar al usuario
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
        AbrRecurso = ObtenerInicialesMayusculas(Recurso)
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
     'Restaurar configuración original
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    Exit Sub
    
ErrorHandler:
    MsgBox "Error al copiar el archivo: " & Err.Description, vbExclamation
' Restaurar configuración original en caso de error
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

Function EvaluarVigencia(vigencia As Date) As String
    Dim wsDashboard As Worksheet
    Dim diasMargen As Integer
    Dim diasDiferencia As Integer
    
    Set wsDashboard = ThisWorkbook.Sheets("DASHBOARD PERSONAL")
    
    ' Leer el valor de la celda G2
    diasMargen = wsDashboard.Range("G2").Value
    
    ' Calcular la diferencia entre la vigencia y la fecha actual
    diasDiferencia = DateDiff("d", Date, vigencia)
    
    ' Evaluar las condiciones
    If diasDiferencia >= diasMargen Then
        EvaluarVigencia = "Vigente"
    ElseIf diasDiferencia > 1 And diasDiferencia < diasMargen Then
        EvaluarVigencia = "Por Vencer"
    Else
        EvaluarVigencia = "Vencido"
    End If
End Function


Sub ActualizarVigenciaDias(rngActiva As Range)
    Dim wsDashboard As Worksheet
    Dim filaActiva As ListRow
    Dim vigencia As Date
    Dim vigenciaDias As Long
    
    Set wsDashboard = ThisWorkbook.Sheets("DASHBOARD PERSONAL")
    
    ' Verificar si la celda activa está en la tabla DashboardPersonal
    If rngActiva Is Nothing Then
        MsgBox "Seleccione una celda dentro de la tabla DashboardPersonal.", vbExclamation
        Exit Sub
    End If
    
    ' Verificar si la fila activa es mayor o igual a la fila 11
    If rngActiva.Row < 11 Then
        Exit Sub
    End If
    
    ' Obtener la fila activa de la tabla DashboardPersonal
    Set filaActiva = wsDashboard.ListObjects("DashboardPersonal").ListRows(rngActiva.Row - wsDashboard.ListObjects("DashboardPersonal").HeaderRowRange.Row)
    
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


Sub GenerarDashboardPersonal(Proyecto As String, TipoRep As Long, Optional Empresa As String = "", Optional ServicioOPuesto As String = "", Optional DNI As String = "")
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsDashboard As Worksheet
    Dim wsPrincipal As Worksheet
    Dim wsPersonal As Worksheet
    Dim wsDocXPuesto As Worksheet
    Dim wsPersonalDocs As Worksheet
    Dim wsDocPorPuesto As Worksheet
    Dim diasMargen As Integer
    Dim link As String
    Dim cell As Range

    Set wsDashboard = ThisWorkbook.Sheets("DASHBOARD PERSONAL")
    Set wsPrincipal = ThisWorkbook.Sheets("PRINCIPAL")
    Set wsPersonal = ThisWorkbook.Sheets("PERSONAL")
    Set wsDocXPuesto = ThisWorkbook.Sheets("MASTER DOCUMENTOS PERSONAL")
    Set wsPersonalDocs = ThisWorkbook.Sheets("PERSONAL DOCUMENTOS")
    Set wsDocPorPuesto = ThisWorkbook.Sheets("DOCUMENTO X PUESTO")

    ' Limpiar la tabla DashboardPersonal
    On Error Resume Next
    wsDashboard.ListObjects("DashboardPersonal").DataBodyRange.Delete
    On Error GoTo 0

    Dim rutaBase As String
    rutaBase = wsPrincipal.Range("B6").Value

    Dim rutaProyecto As String
    Dim rutaRecurso As String
    rutaProyecto = rutaBase & "\" & wsDashboard.Range("C2").Value & "\HABILITACIONES\PERSONAL"
    rutaRecurso = rutaBase & "\RECURSOS\PERSONAL"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim carpetaProyecto As Object
    Set carpetaProyecto = fso.GetFolder(rutaProyecto)

    Dim subcarpeta As Object
    Dim subSubCarpeta As Object

    Dim dniPersonal As String
    Dim fila As ListRow

    ' Cargar datos en arrays para mejorar el rendimiento
    Dim personalData As Variant
    Dim docXPuestoData As Variant
    Dim personalDocsData As Variant
    Dim docPorPuestoData As Variant

    personalData = wsPersonal.Range("A2:E" & wsPersonal.Cells(wsPersonal.Rows.Count, "A").End(xlUp).Row).Value
    docXPuestoData = wsDocXPuesto.Range("A2:D" & wsDocXPuesto.Cells(wsDocXPuesto.Rows.Count, "A").End(xlUp).Row).Value
    personalDocsData = wsPersonalDocs.Range("A2:E" & wsPersonalDocs.Cells(wsPersonalDocs.Rows.Count, "A").End(xlUp).Row).Value
    docPorPuestoData = wsDocPorPuesto.Range("A2:E" & wsDocPorPuesto.Cells(wsDocPorPuesto.Rows.Count, "A").End(xlUp).Row).Value

    Dim personalDict As Object
    Set personalDict = CreateObject("Scripting.Dictionary")
    Dim docXPuestoDict As Object
    Set docXPuestoDict = CreateObject("Scripting.Dictionary")
    Dim personalDocsDict As Object
    Set personalDocsDict = CreateObject("Scripting.Dictionary")
    Dim docPorPuestoDict As Object
    Set docPorPuestoDict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = LBound(personalData, 1) To UBound(personalData, 1)
        ' Combina DNI y Proyecto como la clave en el diccionario
        Dim clave As String
        clave = personalData(i, 1) & "-" & personalData(i, 5) ' Combina DNI y Proyecto como clave única
        ' Almacena los datos en el diccionario usando la clave combinada
        personalDict(clave) = Array(personalData(i, 2), personalData(i, 3), personalData(i, 4), personalData(i, 5))
    Next i

    For i = LBound(docXPuestoData, 1) To UBound(docXPuestoData, 1)
        docXPuestoDict(docXPuestoData(i, 3) & "-" & docXPuestoData(i, 1)) = docXPuestoData(i, 2)
    Next i

    For i = LBound(personalDocsData, 1) To UBound(personalDocsData, 1)
        personalDocsDict(personalDocsData(i, 1) & "-" & personalDocsData(i, 2)) = Array(personalDocsData(i, 3), personalDocsData(i, 4), personalDocsData(i, 5))
    Next i

    For i = LBound(docPorPuestoData, 1) To UBound(docPorPuestoData, 1)
        docPorPuestoDict(docPorPuestoData(i, 4) & "-" & docPorPuestoData(i, 1) & "-" & docPorPuestoData(i, 2)) = docPorPuestoData(i, 3)
    Next i

    diasMargen = wsDashboard.Range("G2").Value

    Dim updateCounter As Long
    updateCounter = 0

    For Each subcarpeta In carpetaProyecto.SubFolders
        dniPersonal = subcarpeta.Name
        
        ' Filtrar por DNI si se ha especificado
        If DNI <> "" And DNI <> dniPersonal Then
            GoTo NextDNI
        End If

        For Each subSubCarpeta In subcarpeta.SubFolders
            ' Crear una nueva fila en la tabla DashboardPersonal
            Set fila = wsDashboard.ListObjects("DashboardPersonal").ListRows.Add
            fila.Range(1, 1).Value = Proyecto

            ' Buscar información del personal
            Dim keyTra As String
            keyTra = dniPersonal & "-" & Proyecto
            If personalDict.Exists(keyTra) Then
                Dim personalInfo As Variant
                personalInfo = personalDict(keyTra)

                ' Filtrar por Empresa y ServicioOPuesto si se han especificado
                If (Empresa <> "" And Empresa <> personalInfo(2)) Or (ServicioOPuesto <> "" And ServicioOPuesto <> personalInfo(1)) Then
                    fila.Delete
                    GoTo NextSubSubCarpeta
                End If

                fila.Range(1, 2).Value = personalInfo(2) ' Empresa
                fila.Range(1, 3).Value = personalInfo(1) ' Servicio o Puesto
                fila.Range(1, 4).Value = dniPersonal ' Doc Identidad
                fila.Range(1, 5).Value = personalInfo(0) ' Apellidos o Nombres
                fila.Range(1, 12).Value = "Sin documento"
            End If

            ' Buscar información del documento
            Dim nombreCarpeta As String
            nombreCarpeta = subSubCarpeta.Name
            Dim keyDoc As String
            keyDoc = nombreCarpeta & "-" & fila.Range(1, 3).Value
            If docXPuestoDict.Exists(keyDoc) Then
                fila.Range(1, 6).Value = docXPuestoDict(keyDoc) ' Documento
                ' Link
                fila.Range(1, 11).Value = "Link"
                fila.Range(1, 11).Hyperlinks.Add Anchor:=fila.Range(1, 11), Address:=rutaRecurso & "\" & dniPersonal & "\" & nombreCarpeta, TextToDisplay:="Link"
            End If

            ' Buscar información de la obligatoriedad del documento
            Dim keyObligatorio As String
            keyObligatorio = Proyecto & "-" & fila.Range(1, 3).Value & "-" & fila.Range(1, 6).Value
            If docPorPuestoDict.Exists(keyObligatorio) Then
                fila.Range(1, 7).Value = docPorPuestoDict(keyObligatorio) ' Obligatorio
            End If

            ' Buscar fechas de emisión y vigencia iterando
            Dim keyPersonalDocs As String
            keyPersonalDocs = dniPersonal & "-" & fila.Range(1, 6).Value
            If personalDocsDict.Exists(keyPersonalDocs) Then
                Dim docsInfo As Variant
                docsInfo = personalDocsDict(keyPersonalDocs)
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
                If fila.Range(1, 7).Value = "SI" Then
                    If vigenciaDias >= diasMargen Then
                        fila.Range(1, 12).Value = "Vigente"
                    ElseIf vigenciaDias >= 1 And vigenciaDias < diasMargen Then
                        fila.Range(1, 12).Value = "Por Vencer"
                    Else
                        fila.Range(1, 12).Value = "Vencido"
                    End If
                Else
                    If vigenciaDias >= diasMargen Then
                        fila.Range(1, 12).Value = "Vigente"
                    ElseIf vigenciaDias >= 1 And vigenciaDias < diasMargen Then
                        fila.Range(1, 12).Value = "Por Vencer"
                    Else
                        fila.Range(1, 12).Value = "Vencido"
                    End If
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
            End If
            
            ' Contador para la actualización periódica
            updateCounter = updateCounter + 1
            If updateCounter Mod 1000 = 0 Then
                Application.ScreenUpdating = True
                Application.ScreenUpdating = False
            End If
            
NextSubSubCarpeta:
        Next subSubCarpeta
NextDNI:
    Next subcarpeta

    ' Restaurar el cálculo automático y la actualización de pantalla
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' Notificar al usuario
    If TipoRep = 1 Then
        MsgBox "Dashboard Personal generado correctamente", vbInformation
    End If
End Sub

Sub GenerarResumenPersonal(Proyecto As String, Optional Empresa As String = "", Optional ServicioOPuesto As String = "", Optional DNI As String = "")
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsDashboard As Worksheet
    Dim wsPrincipal As Worksheet
    Dim wsPersonal As Worksheet
    Dim wsDocXPuesto As Worksheet
    Dim wsPersonalDocs As Worksheet
    Dim wsDocPorPuesto As Worksheet
    Dim wsPersonalResumen As Worksheet
    Dim diasMargen As Integer
    Dim protocolo As String 'Modificación

    Set wsDashboard = ThisWorkbook.Sheets("DASHBOARD PERSONAL")
    Set wsPrincipal = ThisWorkbook.Sheets("PRINCIPAL")
    Set wsPersonal = ThisWorkbook.Sheets("PERSONAL")
    Set wsDocXPuesto = ThisWorkbook.Sheets("MASTER DOCUMENTOS PERSONAL")
    Set wsPersonalDocs = ThisWorkbook.Sheets("PERSONAL DOCUMENTOS")
    Set wsDocPorPuesto = ThisWorkbook.Sheets("DOCUMENTO X PUESTO")
    Set wsPersonalResumen = ThisWorkbook.Sheets("PERSONAL RESUMEN")

    ' Limpiar la tabla PersonalResumen
    On Error Resume Next
    wsPersonalResumen.ListObjects("PersonalResumen").DataBodyRange.Delete
    On Error GoTo 0

    Dim rutaBase As String
    rutaBase = wsPrincipal.Range("B6").Value

    Dim rutaProyecto As String
    Dim rutaRecurso As String
    rutaProyecto = rutaBase & "\" & wsPersonalResumen.Range("C2").Value & "\HABILITACIONES\PERSONAL"
    rutaRecurso = rutaBase & "\RECURSOS\PERSONAL"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim carpetaProyecto As Object
    Set carpetaProyecto = fso.GetFolder(rutaProyecto)

    Dim subcarpeta As Object
    Dim subSubCarpeta As Object

    Dim dniPersonal As String
    Dim fila As ListRow

    ' Cargar datos en arrays para mejorar el rendimiento
    Dim personalData As Variant
    Dim docXPuestoData As Variant
    Dim personalDocsData As Variant
    Dim docPorPuestoData As Variant

    personalData = wsPersonal.Range("A2:G" & wsPersonal.Cells(wsPersonal.Rows.Count, "A").End(xlUp).Row).Value
    docXPuestoData = wsDocXPuesto.Range("A2:D" & wsDocXPuesto.Cells(wsDocXPuesto.Rows.Count, "A").End(xlUp).Row).Value
    personalDocsData = wsPersonalDocs.Range("A2:E" & wsPersonalDocs.Cells(wsPersonalDocs.Rows.Count, "A").End(xlUp).Row).Value
    docPorPuestoData = wsDocPorPuesto.Range("A2:E" & wsDocPorPuesto.Cells(wsDocPorPuesto.Rows.Count, "A").End(xlUp).Row).Value

    Dim personalDict As Object
    Set personalDict = CreateObject("Scripting.Dictionary")
    Dim docXPuestoDict As Object
    Set docXPuestoDict = CreateObject("Scripting.Dictionary")
    Dim docXPuestoDictEmp As Object 'Modificacion
    Set docXPuestoDictEmp = CreateObject("Scripting.Dictionary") 'Modificacion
    Dim personalDocsDict As Object
    Set personalDocsDict = CreateObject("Scripting.Dictionary")
    Dim docPorPuestoDict As Object
    Set docPorPuestoDict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = LBound(personalData, 1) To UBound(personalData, 1)
        ' Combina DNI y Proyecto como la clave en el diccionario
        Dim clave As String
        clave = personalData(i, 1) & "-" & personalData(i, 5) ' Combina DNI y Proyecto como clave única
       ' Almacena los datos en el diccionario usando la clave combinada
        personalDict(clave) = Array(personalData(i, 2), personalData(i, 3), personalData(i, 4), personalData(i, 5), personalData(i, 7))
    Next i

    For i = LBound(docXPuestoData, 1) To UBound(docXPuestoData, 1)
        docXPuestoDict(docXPuestoData(i, 3) & "-" & docXPuestoData(i, 1)) = docXPuestoData(i, 2)
    Next i

    'Modificacion
    For i = LBound(docXPuestoData, 1) To UBound(docXPuestoData, 1)
        docXPuestoDictEmp(docXPuestoData(i, 3) & "-" & docXPuestoData(i, 1)) = Array(docXPuestoData(i, 2), docXPuestoData(i, 4))
    Next i

    For i = LBound(personalDocsData, 1) To UBound(personalDocsData, 1)
        personalDocsDict(personalDocsData(i, 1) & "-" & personalDocsData(i, 2)) = Array(personalDocsData(i, 3), personalDocsData(i, 4), personalDocsData(i, 5))
    Next i

    For i = LBound(docPorPuestoData, 1) To UBound(docPorPuestoData, 1)
        docPorPuestoDict(docPorPuestoData(i, 4) & "-" & docPorPuestoData(i, 1) & "-" & docPorPuestoData(i, 2)) = docPorPuestoData(i, 3)
    Next i

    diasMargen = wsDashboard.Range("G2").Value

    Dim updateCounter As Long
    updateCounter = 0

    Dim habilitados, habilitadosCli, nohabilitados, pendientes, obligatorios As Long
    Dim totalDocs, totalDocsCli As Long

    For Each subcarpeta In carpetaProyecto.SubFolders
        dniPersonal = subcarpeta.Name
        
        ' Filtrar por DNI si se ha especificado
        If DNI <> "" And DNI <> dniPersonal Then
            GoTo NextDNI
        End If

        habilitados = 0
        habilitadosCli = 0 'Modificacion
        pendientes = 0
        nohabilitados = 0
        totalDocs = 0
        totalDocsCli = 0 'Modificacion
        obligatorios = 0
        ' Buscar información del personal
        Dim keyTra As String
        keyTra = dniPersonal & "-" & Proyecto
        If personalDict.Exists(keyTra) Then
            Dim personalInfo As Variant
            personalInfo = personalDict(keyTra)

            ' Filtrar por Empresa y ServicioOPuesto si se han especificado
            If (Empresa <> "" And Empresa <> personalInfo(2)) Or (ServicioOPuesto <> "" And ServicioOPuesto <> personalInfo(1)) Or (personalInfo(3) <> Proyecto) Then
                GoTo NextDNI
            End If

            For Each subSubCarpeta In subcarpeta.SubFolders
                

                ' Buscar información del documento
                Dim nombreCarpeta As String
                nombreCarpeta = subSubCarpeta.Name
                Dim keyDoc As String
                keyDoc = nombreCarpeta & "-" & personalInfo(1)
                If docXPuestoDict.Exists(keyDoc) Then
                    Dim keyObligatorio As String
                    keyObligatorio = Proyecto & "-" & personalInfo(1) & "-" & docXPuestoDict(keyDoc)
                    docObligatorio = docPorPuestoDict(keyObligatorio)
                    protocolo = UCase(docXPuestoDictEmp(keyDoc)(1)) 'Modificacion
                    If docObligatorio = "SI" Then
                        ' Contar documentos habilitados
                        'Modificacion
                        If protocolo = "TEMA" Then
                            totalDocs = totalDocs + 1
                        Else
                            totalDocsCli = totalDocsCli + 1
                        End If
                    End If
                    ' Buscar fechas de emisión y vigencia iterando
                    Dim keyPersonalDocs As String
                    keyPersonalDocs = dniPersonal & "-" & docXPuestoDict(keyDoc)
                    If personalDocsDict.Exists(keyPersonalDocs) Then
                        Dim docsInfo As Variant
                        docsInfo = personalDocsDict(keyPersonalDocs)
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
            Set fila = wsPersonalResumen.ListObjects("PersonalResumen").ListRows.Add
            fila.Range(1, 1).Value = Proyecto
            fila.Range(1, 2).Value = personalInfo(2) ' Empresa
            fila.Range(1, 3).Value = personalInfo(1) ' Servicio o Puesto
            fila.Range(1, 4).Value = "'" & dniPersonal ' Doc Identidad
            fila.Range(1, 5).Value = personalInfo(0) ' Apellidos y Nombres
            fila.Range(1, 6).Value = habilitados ' Documentos Habilitados Tema
            fila.Range(1, 7).Value = totalDocs ' Total Documentos Tema
            'Modificacion
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
            fila.Range(1, 11).Value = personalInfo(4) ' Situación
        End If

NextDNI:
    Next subcarpeta

    ' Restaurar el cálculo automático y la actualización de pantalla
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' Notificar al usuario
    MsgBox "Resumen Personal generado correctamente", vbInformation
End Sub

Sub ActivaResumenPersonal()
    ' Activar la hoja "PERSONAL RESUMEN"
    ThisWorkbook.Sheets("PERSONAL RESUMEN").Activate
End Sub


