Attribute VB_Name = "Módulo3"
Sub EliminarEspaciosIzquierda()
    Dim celda As Range
    Dim lineas() As String
    Dim i As Integer
    
    ' Recorre cada celda en la selección
    For Each celda In Selection
        ' Divide el contenido de la celda en líneas
        lineas = Split(celda.Value, vbCrLf)
        ' Recorre cada línea y elimina los espacios en blanco a la izquierda
        For i = LBound(lineas) To UBound(lineas)
            lineas(i) = Trim(Replace(lineas(i), Chr(9), ""))
        Next i
        ' Actualiza el contenido de la celda con el texto limpio
        celda.Value = Join(lineas, vbCrLf)
    Next celda
End Sub

