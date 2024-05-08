Attribute VB_Name = "M�dulo3"
Sub EliminarEspaciosIzquierda()
    Dim celda As Range
    Dim lineas() As String
    Dim i As Integer
    
    ' Recorre cada celda en la selecci�n
    For Each celda In Selection
        ' Divide el contenido de la celda en l�neas
        lineas = Split(celda.Value, vbCrLf)
        ' Recorre cada l�nea y elimina los espacios en blanco a la izquierda
        For i = LBound(lineas) To UBound(lineas)
            lineas(i) = Trim(Replace(lineas(i), Chr(9), ""))
        Next i
        ' Actualiza el contenido de la celda con el texto limpio
        celda.Value = Join(lineas, vbCrLf)
    Next celda
End Sub

