Attribute VB_Name = "M�dulo1"
Sub EliminarComaSobrante()

    Dim celda As Range
    
    ' Recorre cada celda en la selecci�n
    For Each celda In Selection
        ' Verifica si la celda no est� vac�a y contiene texto
        If Not IsEmpty(celda.Value) And Len(celda.Value) > 0 Then
            ' Verifica y elimina la coma al final de la celda
            If Right(celda.Value, 1) = "," Then
                celda.Value = Left(celda.Value, Len(celda.Value) - 1)
            End If
            ' Verifica y elimina la coma al inicio de la celda
            If Left(celda.Value, 1) = "," Then
                celda.Value = Mid(celda.Value, 2)
            End If
        End If
    Next celda

End Sub

