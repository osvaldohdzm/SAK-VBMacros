Attribute VB_Name = "Módulo1"
Sub LimpiarCeldas()

    Dim celda As Range
    Dim valores() As String
    Dim nuevoValor As String
    
    ' Recorre cada celda seleccionada
    For Each celda In Selection
        ' Verifica si la celda no está vacía
        If Not IsEmpty(celda.Value) Then
            ' Divide los valores por saltos de línea
            valores = Split(celda.Value, Chr(10))
            
            ' Obtiene solo el primer valor
            nuevoValor = valores(0)
            
            ' Asigna el nuevo valor a la celda
            celda.Value = nuevoValor
        End If
    Next celda

End Sub

