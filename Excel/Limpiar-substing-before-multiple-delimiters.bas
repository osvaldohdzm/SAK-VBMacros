Attribute VB_Name = "M�dulo3"
Sub SepararPorMultiplesDelimitadores()
    Dim celda As Range
    Dim valor As String
    
    For Each celda In Selection
        If celda.Value <> "" Then
            valor = celda.Value
            ' Buscar el primer salto de l�nea, dos puntos, coma o espacio y conservar la parte antes de �l
            valor = Split(Split(Split(Split(valor, vbLf)(0), ":")(0), ",")(0), " ")(0)
            celda.Value = valor
        End If
    Next celda
End Sub

