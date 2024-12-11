Attribute VB_Name = "Módulo1"
Sub ExportarTabla()
    Dim celdaActual As Range
    Dim tabla As ListObject
    Dim rutaArchivo As String
    Dim nombreArchivo As String
    Dim rutaArchivoOriginal As String

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
        
        ' Definir la ruta del archivo (misma ruta que el archivo de Excel)
        rutaArchivo = ThisWorkbook.Path & "\" & nombreArchivo
        
        ' Guardar la ruta original del archivo con su extensión original
        rutaArchivoOriginal = ThisWorkbook.FullName
        
        ' Verificar si la tabla tiene más de una columna
        If tabla.ListColumns.Count > 1 Then
            ' Exportar a CSV si tiene más de una columna
            ExportarTablaCSV rutaArchivo & ".csv", tabla
        Else
            ' Exportar a TXT si tiene una sola columna
            ExportarTablaTXT rutaArchivo & ".txt", tabla
        End If
        
        MsgBox "Archivo exportado con éxito: " & rutaArchivo
        
        ' Guardar el archivo original nuevamente
        Application.DisplayAlerts = False
        ThisWorkbook.SaveAs Filename:=rutaArchivoOriginal, FileFormat:=ThisWorkbook.FileFormat
        Application.DisplayAlerts = True
        
    Else
        MsgBox "La celda seleccionada no está dentro de una tabla."
    End If
End Sub

' Función para exportar contenido de una tabla a formato CSV
Sub ExportarTablaCSV(rutaArchivo As String, tabla As ListObject)
    ' Guardar la tabla como CSV
    Application.DisplayAlerts = False
    tabla.Range.Copy
    tabla.Parent.Cells.PasteSpecial xlPasteValues
    ActiveWorkbook.SaveAs Filename:=rutaArchivo, FileFormat:=xlCSV, CreateBackup:=False
    Application.DisplayAlerts = True
End Sub

' Función para exportar contenido de una tabla a formato TXT
Sub ExportarTablaTXT(rutaArchivo As String, tabla As ListObject)
    ' Guardar la tabla como archivo de texto
    Application.DisplayAlerts = False
    tabla.Range.Copy
    tabla.Parent.Cells.PasteSpecial xlPasteValues
    ActiveWorkbook.SaveAs Filename:=rutaArchivo, FileFormat:=xlText, CreateBackup:=False
    Application.DisplayAlerts = True
End Sub

