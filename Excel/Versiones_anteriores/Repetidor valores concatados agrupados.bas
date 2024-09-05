Attribute VB_Name = "Módulo1"
Sub repetir_concatacion_agrupada()
    Dim selectedRange As Range
    Dim cell As Range
    Dim concatenatedText As String
    
    ' Verificar que hay celdas seleccionadas
    If Selection.Cells.Count > 1 Then
        ' Recorrer cada celda seleccionada
        For Each selectedRange In Selection.Areas
            concatenatedText = ""
            ' Concatenar los valores con salto de línea
            For Each cell In selectedRange.Cells
                concatenatedText = concatenatedText & cell.Value & vbCrLf
            Next cell
            ' Eliminar el último salto de línea adicional
            concatenatedText = Left(concatenatedText, Len(concatenatedText) - 2)
            
            ' Asignar el valor concatenado a cada celda de la selección
            For Each cell In selectedRange.Cells
                cell.Value = concatenatedText
            Next cell
        Next selectedRange
    Else
        MsgBox "Selecciona más de una celda para ejecutar esta macro."
    End If
End Sub

