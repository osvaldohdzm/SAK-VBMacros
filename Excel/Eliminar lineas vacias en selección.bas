Attribute VB_Name = "Module1"
Sub EliminarLineasVaciasEnCeldasSeleccionadas()
    Dim celda As Range
    Dim lineas As Variant
    Dim i As Integer

    ' Iterar sobre cada celda seleccionada
    For Each celda In Selection
        ' Verificar si la celda tiene texto
        If Not IsEmpty(celda.Value) Then
            ' Reemplazar diferentes saltos de línea con vbLf
            Dim contenido As String
            contenido = Replace(Replace(Replace(celda.Value, vbCrLf, vbLf), vbCr, vbLf), vbLf & vbLf, vbLf)
            
            ' Si el contenido comienza con vbLf, quitarlo
            If Left(contenido, 1) = vbLf Then
                contenido = Mid(contenido, 2)
            End If
            
            ' Si el contenido termina con vbLf, quitarlo
            If Right(contenido, 1) = vbLf Then
                contenido = Left(contenido, Len(contenido) - 1)
            End If
            
            ' Dividir el contenido de la celda en un array de líneas
            lineas = Split(contenido, vbLf)
            
            ' Crear un nuevo array para almacenar las líneas no vacías
            Dim lineasSinVacias() As String
            ReDim lineasSinVacias(0 To UBound(lineas))
            Dim idx As Integer
            idx = 0
            
            ' Iterar sobre cada línea del array
            For i = LBound(lineas) To UBound(lineas)
                ' Verificar si la línea está vacía y no agregarla al nuevo array
                If Trim(lineas(i)) <> "" Then
                    lineasSinVacias(idx) = lineas(i)
                    idx = idx + 1
                End If
            Next i
            
            ' Redimensionar el array resultante
            ReDim Preserve lineasSinVacias(0 To idx - 1)
            
            ' Unir el array de líneas de nuevo en una cadena y asignarlo a la celda
            celda.Value = Join(lineasSinVacias, vbLf)
        End If
    Next celda
End Sub

