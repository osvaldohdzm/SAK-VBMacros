Attribute VB_Name = "Módulo3"
Sub SepararPorMultiplesDelimitadores()
    Dim celda As Range
    Dim valor As String
    
    For Each celda In Selection
        If celda.Value <> "" Then
            valor = celda.Value
            ' Buscar el primer salto de línea, dos puntos, coma o espacio y conservar la parte antes de él
            valor = Split(Split(Split(Split(valor, vbLf)(0), ":")(0), ",")(0), " ")(0)
            celda.Value = valor
        End If
    Next celda
End Sub

