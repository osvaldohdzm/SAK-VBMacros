Attribute VB_Name = "Módulo1"
Sub ReemplazarPalabras()
    Dim c As Range
    Dim valorActual As String
    
    For Each c In Selection
        valorActual = Trim(UCase(c.Value)) ' Convertimos a mayúsculas y eliminamos espacios adicionales
        
        Select Case valorActual
            Case "0", "NONE", "INFORMATIVA", "INFO"
                c.Value = "INFORMATIVA"
            Case "1", "BAJA", "BAJO", "LOW"
                c.Value = "BAJA"
            Case "2", "BAJA", "BAJO", "LOW"
                c.Value = "BAJA"
            Case "3", "BAJA", "BAJO", "LOW"
                c.Value = "BAJA"
            Case "4", "BAJA", "BAJO", "LOW"
                c.Value = "BAJA"
            Case "5", "MEDIA", "MEDIO", "MEDIUM"
                c.Value = "MEDIA"
            Case "6", "MEDIA", "MEDIO", "MEDIUM"
                c.Value = "MEDIA"
            Case "7", "ALTO", "HIGH"
                c.Value = "ALTA"
            Case "8", "ALTA", "HIGH"
                c.Value = "ALTA"
            Case "9", "CRÍTICA", "CRITICAL", "CRÍTICO"
                c.Value = "CRÍTICA"
            Case "10", "CRÍTICA", "CRITICAL", "CRÍTICO"
                c.Value = "CRÍTICA"
            ' Mantener el contenido actual si no coincide con las palabras a reemplazar
            Case Else
                ' No hacer nada
        End Select
    Next c
End Sub

