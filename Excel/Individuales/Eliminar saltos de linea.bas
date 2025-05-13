Attribute VB_Name = "Module2"
Sub EliminarSaltosDeLinea()

    Dim Celda As Range
    Dim Texto As String
    Dim NuevoTexto As String
    
    ' Itera a través de las celdas seleccionadas en la hoja activa
    For Each Celda In Selection
        If Not Celda.HasFormula Then ' Ignora celdas con fórmulas
            Texto = Celda.Value
            
            ' Reemplazar diferentes tipos de saltos de línea y retornos de carro
            NuevoTexto = Replace(Texto, vbCrLf, " ")   ' Salto de línea + retorno de carro
            NuevoTexto = Replace(NuevoTexto, vbCr, " ") ' Retorno de carro
            NuevoTexto = Replace(NuevoTexto, vbLf, " ") ' Salto de línea
            
            Celda.Value = NuevoTexto ' Asigna el nuevo valor a la celda
        End If
    Next Celda

End Sub

