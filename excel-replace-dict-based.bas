Attribute VB_Name = "Módulo1"
Sub ReemplazarPalabras()
    Dim c As Range
    Dim valorActual As String
    
    For Each c In Selection
        valorActual = c.Value
        
        Select Case valorActual
            Case "Moderate"
                c.Value = "MEDIO"
            Case "Critical"
                c.Value = "CRÍTICO"
            Case "Important"
                c.Value = "ALTO"
            Case "Low"
                c.Value = "BAJO"
            Case "Info"
                c.Value = "INFO"
            Case "High"
                c.Value = "ALTO"
            Case "Medium"
                c.Value = "MEDIO"
            Case "Low"
                c.Value = "BAJO"
            ' Mantener el contenido actual si no coincide con las palabras a reemplazar
            Case Else
                ' No hacer nada
        End Select
    Next c
End Sub

