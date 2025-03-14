Attribute VB_Name = "ModiifcaAccesosDirectos"
Sub ActualizarAccesosDirectos()
    Dim objShell As Object
    Dim objFSO As Object
    Dim carpetaActual As String

    ' Crear objetos de Shell y FileSystemObject
    Set objShell = CreateObject("WScript.Shell")
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' Obtener la carpeta donde se ejecuta el script
    carpetaActual = "C:\Users\lhoyos\Tema Litoclean\GESPRO - Documentos\General\10706 HABILITACIONES\General"

    ' Llamar a la función recursiva para buscar accesos directos en subcarpetas
    BuscarYActualizarLNK objFSO.GetFolder(carpetaActual), objShell, carpetaActual

    ' Liberar objetos
    Set objShell = Nothing
    Set objFSO = Nothing

    MsgBox "Accesos directos actualizados en todas las carpetas.", vbInformation
End Sub

Sub BuscarYActualizarLNK(carpeta As Object, objShell As Object, carpetaBase As String)
    Dim archivo As Object
    Dim subcarpeta As Object
    Dim objShortcut As Object
    Dim nuevaRuta As String

    ' Recorrer todos los archivos en la carpeta
    For Each archivo In carpeta.Files
        ' Verificar si es un acceso directo (.lnk)
        If LCase(Right(archivo.Name, 4)) = ".lnk" Then
            ' Cargar acceso directo
            Set objShortcut = objShell.CreateShortcut(archivo.Path)

            ' Reemplazar la parte del path antiguo con la nueva ubicación
            nuevaRuta = Replace(objShortcut.targetPath, "H:\", "C:\Users\lhoyos\Tema Litoclean\GESPRO - Documentos\General\10706 HABILITACIONES\General\")

            ' Actualizar la ruta si es necesario
            If objShortcut.targetPath <> nuevaRuta Then
                objShortcut.targetPath = nuevaRuta
                objShortcut.Save
            End If
        End If
    Next archivo

    ' Recorrer todas las subcarpetas y aplicar la función recursivamente
    For Each subcarpeta In carpeta.SubFolders
        BuscarYActualizarLNK subcarpeta, objShell, carpetaBase
    Next subcarpeta
End Sub
Sub LeerAccesosDirectosSubcarpetas()
    Dim fso As Object
    Dim carpetaPrincipal As Object
    Dim wsh As Object
    Dim ws As Worksheet
    Dim fila As Integer
    Dim carpetaRuta As String

    ' Definir la carpeta donde están los accesos directos
    carpetaRuta = "C:\Users\lhoyos\Tema Litoclean\GESPRO - Documentos\General\10706 HABILITACIONES\General" ' <-- Cambia esto por la ruta real

    ' Crear objetos
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set carpetaPrincipal = fso.GetFolder(carpetaRuta)
    Set wsh = CreateObject("WScript.Shell")

    ' Definir la hoja activa o crear una nueva hoja
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets("Accesos_Directos")
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Sheets.Add
        ws.Name = "Accesos_Directos"
    End If
    On Error GoTo 0

    ' Limpiar la hoja antes de escribir
    ws.Cells.Clear
    
    ' Escribir encabezados
    ws.Cells(1, 1).Value = "Nombre del Acceso Directo"
    ws.Cells(1, 2).Value = "Ruta Real del Archivo"
    ws.Cells(1, 3).Value = "Ubicación (Carpeta donde se encontró)"
    ws.Cells(1, 4).Value = "Fecha de Creación"

    ' Formato en negrita
    ws.Rows(1).Font.Bold = True

    fila = 2 ' Empezar en la segunda fila para los datos

    ' Llamar a la función recursiva para recorrer todas las carpetas
    ProcesarCarpeta carpetaPrincipal, ws, fila, fso, wsh

    ' Ajustar el ancho de las columnas
    ws.Columns("A:D").AutoFit

    ' Liberar objetos
    Set fso = Nothing
    Set wsh = Nothing
    Set carpetaPrincipal = Nothing

    MsgBox "Proceso completado. Datos guardados en la hoja 'Accesos_Directos'.", vbInformation
End Sub

' Función recursiva para procesar cada carpeta y subcarpeta
Sub ProcesarCarpeta(carpeta As Object, ws As Worksheet, ByRef fila As Integer, fso As Object, wsh As Object)
    Dim archivo As Object
    Dim subcarpeta As Object
    Dim accesoDirecto As Object
    Dim rutaReal As String

    ' Recorrer todos los archivos en la carpeta actual
    For Each archivo In carpeta.Files
        If LCase(fso.GetExtensionName(archivo.Name)) = "lnk" Then
            ' Obtener el acceso directo
            Set accesoDirecto = wsh.CreateShortcut(archivo.Path)
            rutaReal = accesoDirecto.targetPath ' Obtener la ruta real

            ' Escribir en la hoja de Excel
            ws.Cells(fila, 1).Value = archivo.Name
            ws.Cells(fila, 2).Value = rutaReal
            ws.Cells(fila, 3).Value = carpeta.Path ' Carpeta donde se encontró
            ws.Cells(fila, 4).Value = archivo.DateCreated ' Fecha de creación

            fila = fila + 1
            
            ' Liberar objeto
            Set accesoDirecto = Nothing
        End If
    Next archivo

    ' Recorrer todas las subcarpetas de la carpeta actual
    For Each subcarpeta In carpeta.SubFolders
        ProcesarCarpeta subcarpeta, ws, fila, fso, wsh ' Llamada recursiva
    Next subcarpeta
End Sub

