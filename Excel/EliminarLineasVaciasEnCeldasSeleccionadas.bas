Attribute VB_Name = "M�dulo1"
Sub EliminarLineasVaciasEnCeldasSeleccionadas()
    Dim celda As Range
    Dim lineas As Variant
    Dim i As Integer

    ' Iterar sobre cada celda seleccionada
    For Each celda In Selection
        ' Verificar si la celda tiene texto
        If Not IsEmpty(celda.Value) Then
            ' Reemplazar diferentes saltos de l�nea con vbLf
            Dim contenido As String
            contenido = Replace(Replace(Replace(celda.Value, vbCrLf, vbLf), vbCr, vbLf), vbLf & vbLf, vbLf)
            
            ' Si el contenido termina con vbLf, quitarlo
            If Right(contenido, 1) = vbLf Then
                contenido = Left(contenido, Len(contenido) - 1)
            End If
            
            ' Dividir el contenido de la celda en un array de l�neas
            lineas = Split(contenido, vbLf)
            
            ' Iterar sobre cada l�nea del array
            For i = LBound(lineas) To UBound(lineas)
                ' Verificar si la l�nea est� vac�a y eliminarla
                If Trim(lineas(i)) = "" Then
                    lineas(i) = vbNullString
                End If
            Next i
            
            ' Unir el array de l�neas de nuevo en una cadena y asignarlo a la celda
            celda.Value = Join(lineas, vbLf)
        End If
    Next celda
End Sub

