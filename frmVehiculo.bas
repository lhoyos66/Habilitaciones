VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVehiculo 
   Caption         =   "Vehículos - Maquinarias"
   ClientHeight    =   3480
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   8180
   OleObjectBlob   =   "frmVehiculo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmVehiculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nombreHojaActiva As String

Private Sub cmdSeleccionar_Click()
    Unload Me
End Sub

Private Sub ListBox1_Click()
 Dim seleccion As String
    If ListBox1.ListIndex <> -1 Then ' Asegúrate de que se haya seleccionado algo
        seleccion = ListBox1.Value
        Dim codigo As String
        codigo = Left(seleccion, InStr(seleccion, " - ") - 1)
        
        ' Retorna el código a la celda activa
        With ActiveCell
            .Value = "'" & codigo ' Asegura que Excel trate el valor como texto
        End With
        Me.Hide
    End If
End Sub

Private Sub TextBox1_Change()
    FiltrarLista TextBox1.Text
End Sub

Public Sub LlenarListBoxConDatos(nombreHoja As String)
    Dim hoja As Worksheet
    Set hoja = ThisWorkbook.Sheets(nombreHoja)
    
    ' Asegúrate de que la hoja tiene una tabla llamada "Vehiculo"
    Dim tbl As ListObject
    Set tbl = hoja.ListObjects("Vehiculo")
    
    ' Limpiar el ListBox
    Me.ListBox1.Clear
    
    ' Llenar el ListBox con los datos de la tabla
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Me.ListBox1.AddItem tbl.DataBodyRange(i, 1).Value & " - " & tbl.DataBodyRange(i, 2).Value
    Next i
End Sub

Public Sub FiltrarLista(criterio As String)
    Dim hoja As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim nombreHoja As String

    nombreHoja = "MASTER VEHICULOS" ' Nombre de la hoja que contiene la tabla de proveedores

    ' Control de errores
    On Error GoTo ErrorHandler

    ' Verifica si la hoja existe
    Set hoja = ThisWorkbook.Sheets(nombreHoja)

    ' Verifica si la tabla existe
    Set tbl = hoja.ListObjects("Vehiculo")

    ' Limpia los resultados anteriores
    Me.ListBox1.Clear

    ' Convertir el criterio a mayúsculas
    criterio = UCase(criterio)

    ' Filtra y añade a la lista
    For i = 1 To tbl.ListRows.Count ' Recorre todas las filas de la tabla
        If InStr(1, UCase(tbl.DataBodyRange(i, 1).Value), criterio, vbTextCompare) > 0 Or _
           InStr(1, UCase(tbl.DataBodyRange(i, 2).Value), criterio, vbTextCompare) > 0 Then
            Me.ListBox1.AddItem tbl.DataBodyRange(i, 1).Value & " - " & tbl.DataBodyRange(i, 2).Value
        End If
    Next i

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub



