Attribute VB_Name = "Módulo1"
Sub ReemplazarAntesComa()
    Dim celda As Range
    For Each celda In Selection
        If celda.Value <> "" Then
            celda.Value = Split(celda.Value, ",")(0)
        End If
    Next celda
End Sub

