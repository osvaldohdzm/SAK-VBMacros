Attribute VB_Name = "Módulo1"
Sub ReemplazarAntesSaltoDeLinea()
    Dim celda As Range
    For Each celda In Selection
        If celda.Value <> "" Then
            Dim partes() As String
            partes = Split(celda.Value, vbLf)
            If UBound(partes) >= 1 Then
                celda.Value = partes(0)
            End If
        End If
    Next celda
End Sub

