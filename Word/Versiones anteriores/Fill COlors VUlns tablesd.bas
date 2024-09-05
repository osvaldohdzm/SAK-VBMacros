Attribute VB_Name = "NewMacros"
Sub AplicarFormatoACeldas()
    Dim tabla As table
    Dim celda As Cell

    ' Iterar sobre cada tabla en el documento
    For Each tabla In ActiveDocument.Tables
        ' Verificar si la tabla tiene al menos dos columnas y una fila
        If tabla.Columns.count >= 2 And tabla.Rows.count >= 1 Then
            ' Obtener la celda en la primera fila y segunda columna
            Set celda = tabla.Cell(1, 2)
            
            ' Verificar si el contenido contiene "Bajo" (sin importar mayúsculas o minúsculas)
            If InStr(1, UCase(Trim(celda.Range.Text)), "BAJO", vbTextCompare) > 0 Then
                ' Seleccionar la celda
                celda.Range.Cells(1).Select
                
                ' Aplicar formato a la selección
                With Selection
                    .Shading.BackgroundPatternColor = RGB(0, 176, 80) ' Verde (#00B050)
                    .Font.Color = RGB(255, 255, 255) ' Blanco
                End With
            End If
        End If
    Next tabla
End Sub

