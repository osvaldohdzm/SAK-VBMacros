Attribute VB_Name = "M�dulo1"
Sub repetir_concatacion_agrupada()
    Dim selectedRange As Range
    Dim cell As Range
    Dim concatenatedText As String
    
    ' Verificar que hay celdas seleccionadas
    If Selection.Cells.Count > 1 Then
        ' Recorrer cada celda seleccionada
        For Each selectedRange In Selection.Areas
            concatenatedText = ""
            ' Concatenar los valores con salto de l�nea
            For Each cell In selectedRange.Cells
                concatenatedText = concatenatedText & cell.Value & vbCrLf
            Next cell
            ' Eliminar el �ltimo salto de l�nea adicional
            concatenatedText = Left(concatenatedText, Len(concatenatedText) - 2)
            
            ' Asignar el valor concatenado a cada celda de la selecci�n
            For Each cell In selectedRange.Cells
                cell.Value = concatenatedText
            Next cell
        Next selectedRange
    Else
        MsgBox "Selecciona m�s de una celda para ejecutar esta macro."
    End If
End Sub

