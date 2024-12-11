Attribute VB_Name = "Module2"
Sub EliminarSaltosDeLinea()

    Dim Celda As Range
    Dim Texto As String
    Dim NuevoTexto As String
    
    ' Itera a trav�s de las celdas seleccionadas en la hoja activa
    For Each Celda In Selection
        If Not Celda.HasFormula Then ' Ignora celdas con f�rmulas
            Texto = Celda.Value
            
            ' Reemplazar diferentes tipos de saltos de l�nea y retornos de carro
            NuevoTexto = Replace(Texto, vbCrLf, " ")   ' Salto de l�nea + retorno de carro
            NuevoTexto = Replace(NuevoTexto, vbCr, " ") ' Retorno de carro
            NuevoTexto = Replace(NuevoTexto, vbLf, " ") ' Salto de l�nea
            
            Celda.Value = NuevoTexto ' Asigna el nuevo valor a la celda
        End If
    Next Celda

End Sub

