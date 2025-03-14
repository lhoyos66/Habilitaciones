Attribute VB_Name = "OpcionesPrincipales"
Dim ocultar As Boolean
Sub OcultarOMostrarHojas()
        Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo ErrorHandler
    
    If ocultar Then
        OcultarHojas
        ThisWorkbook.Worksheets("PRINCIPAL").Activate
        ThisWorkbook.Worksheets("PRINCIPAL").Shapes("OcultarHojasDeApoyo").TextFrame2.TextRange.Text = "Mostrar Hojas de Apoyo"
    Else
        MostrarHojas
        ThisWorkbook.Worksheets("PRINCIPAL").Activate
        ThisWorkbook.Worksheets("PRINCIPAL").Shapes("OcultarHojasDeApoyo").TextFrame2.TextRange.Text = "Ocultar Hojas de Apoyo"
    End If
    
    ocultar = Not ocultar
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub OcultarHojas()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "PRINCIPAL" And ws.Name <> "GESTION INTEGRAL CAMPO" And ws.Name <> "GESTION INTEGRAL VEHICULO CAMPO" And ws.Name <> "GESTION INTEGRAL EMBARCACIONES" And ws.Name <> "DASHBOARD PERSONAL" And ws.Name <> "DASHBOARD VEHICULO" And ws.Name <> "DASHBOARD EMBARCACIONES" And ws.Name <> "PERSONAL RESUMEN" And ws.Name <> "VEHICULO RESUMEN" And ws.Name <> "EMBARCACION RESUMEN" Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
End Sub

Sub MostrarHojas()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVeryHidden Then
            ws.Visible = xlSheetVisible
        End If
    Next ws
End Sub
Sub CambiarAConfiguracionProyecto()
    Sheets("CONFIGURACIÓN PROYECTO").Activate
End Sub

Sub CambiarAHojaPanelPersonal()
    Sheets("PERSONAL RESUMEN").Activate
End Sub


Sub CambiarAHojaPanelVehiculos()
    Sheets("VEHICULO RESUMEN").Activate
End Sub

Sub CambiarAHojaPanelEmbarcaciones()
    Sheets("EMBARCACION RESUMEN").Activate
End Sub


Sub SeleccionarCarpeta()
    Dim fd As FileDialog
    Dim folderPath As String
    
    ' Crear una ventana de diálogo para seleccionar una carpeta
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    ' Configurar propiedades de la ventana de diálogo
    With fd
        .Title = "Seleccionar Carpeta"
        .AllowMultiSelect = False
        If .Show = -1 Then ' Si el usuario selecciona una carpeta
            folderPath = .SelectedItems(1)
            ' Escribir la ruta de la carpeta en la celda B6 de la hoja Principal
            ThisWorkbook.Sheets("PRINCIPAL").Range("B6").Value = folderPath
        Else ' Si el usuario cancela la selección
            MsgBox "No se seleccionó ninguna carpeta", vbExclamation
        End If
    End With
    
    ' Limpiar la variable de la ventana de diálogo
    Set fd = Nothing
End Sub


Sub SeleccionarArchivoCentral()
    Dim fd As FileDialog
    Dim filePath As String
    
    ' Crear una ventana de diálogo para seleccionar un archivo
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    ' Configurar propiedades de la ventana de diálogo
    With fd
        .Title = "Seleccionar Archivo"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Todos los archivos", "*.*"
        If .Show = -1 Then ' Si el usuario selecciona un archivo
            filePath = .SelectedItems(1)
            ' Escribir la ruta del archivo en la celda B8 de la hoja PRINCIPAL
            ThisWorkbook.Sheets("PRINCIPAL").Range("B8").Value = filePath
        Else ' Si el usuario cancela la selección
            MsgBox "No se seleccionó ningún archivo", vbExclamation
        End If
    End With
    
    ' Limpiar la variable de la ventana de diálogo
    Set fd = Nothing
End Sub


Sub CambiarAHojaPrincipal()
    Sheets("PRINCIPAL").Activate
End Sub
