Attribute VB_Name = "ExportaModulos"
Sub ExportarModulosVBA()
    Dim vbComp As Object
    Dim exportPath As String
    Dim fd As FileDialog
    
    'Crear diálogo de selección de carpeta
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Seleccione carpeta para exportar módulos"
        .AllowMultiSelect = False
        If .Show = -1 Then
            exportPath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    If Right(exportPath, 1) <> "\" Then exportPath = exportPath & "\"
    
    'Exportar módulos
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1, 2, 3 'Módulos, clases y formularios
                vbComp.Export exportPath & vbComp.Name & ".bas"
        End Select
    Next vbComp
    
    MsgBox "Módulos exportados a " & exportPath, vbInformation
End Sub

