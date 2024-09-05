Attribute VB_Name = "ExcelMacrosGeneral"
Sub ExportarTabla()
    Dim celdaActual As Range
    Dim tabla As ListObject
    Dim rutaArchivo As String
    Dim nombreArchivo As String
    Dim carpetaDestino As String
    Dim nuevoLibro As Workbook
    Dim nuevaHoja As Worksheet
    Dim archivoGuardado As Variant
    
    ' Obtener la celda actualmente seleccionada
    Set celdaActual = ActiveCell
    
    ' Verificar si la celda seleccionada está dentro de una tabla
    On Error Resume Next
    Set tabla = celdaActual.ListObject
    On Error GoTo 0
    
    ' Si la celda está dentro de una tabla, procedemos
    If Not tabla Is Nothing Then
        ' Obtener el nombre de la tabla
        nombreArchivo = tabla.Name
        
        ' Mostrar un cuadro de diálogo para seleccionar la carpeta de destino
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Selecciona la carpeta para guardar el archivo"
            .Show
            If .SelectedItems.Count > 0 Then
                carpetaDestino = .SelectedItems(1)
            Else
                MsgBox "No se seleccionó ninguna carpeta. La exportación se ha cancelado.", vbExclamation
                Exit Sub
            End If
        End With
        
        ' Definir la ruta del archivo
        rutaArchivo = carpetaDestino & "\" & nombreArchivo & ".csv"
        
        ' Crear una nueva instancia de Excel
        Set nuevoLibro = Workbooks.Add
        Set nuevaHoja = nuevoLibro.Sheets(1)
        
        ' Copiar la tabla a la nueva hoja
        tabla.Range.Copy
        nuevaHoja.Cells.PasteSpecial xlPasteValues
        
        ' Guardar la nueva hoja como archivo CSV
        Application.DisplayAlerts = False
        nuevoLibro.SaveAs Filename:=rutaArchivo, FileFormat:=xlCSV, CreateBackup:=False
        Application.DisplayAlerts = True
        
        ' Cerrar la nueva instancia de Excel sin guardar cambios
        nuevoLibro.Close SaveChanges:=False
        
        MsgBox "Archivo exportado con éxito: " & rutaArchivo
        
        ' Regresar a la hoja original
        ThisWorkbook.Sheets(1).Activate
        
    Else
        MsgBox "La celda seleccionada no está dentro de una tabla."
    End If
End Sub



Sub Lowercase()
 For Each cell In Selection
        If Not cell.HasFormula Then
            cell.Value = LCase(cell.Value)
        End If
    Next cell
End Sub

Sub AjustarAlturaFilasEnTodasLasHojasDelLibroActivo()
    Dim sh As Worksheet
    
    ' Recorre todas las hojas en el libro activo
    For Each sh In ActiveWorkbook.Worksheets
        ' Ajusta la altura de todas las filas en la hoja actual
        sh.Rows.RowHeight = 15
    Next sh
    
    ' Muestra un mensaje indicando que la operación se completó
    MsgBox "Todas las filas en todas las hojas del libro activo se han ajustado a una altura de 15."
End Sub


