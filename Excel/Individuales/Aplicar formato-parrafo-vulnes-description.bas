Attribute VB_Name = "Módulo1"
Sub ReemplazarSaltosDeLineaConEspacios()
    Dim Rango As Range
    Dim Celda As Range
    Dim Texto As String
    
    ' Verificar si hay celdas seleccionadas
    If Selection.Cells.Count = 0 Then
        MsgBox "No has seleccionado ninguna celda.", vbExclamation
        Exit Sub
    End If
    
    ' Recorrer las celdas seleccionadas
    For Each Rango In Selection
        If Rango.HasFormula = False Then
            Texto = Rango.Value
            Texto = Replace(Texto, vbLf & vbLf, "\n")
            Texto = Replace(Texto, vbLf, " ")
            Texto = Replace(Texto, "\n", vbLf)
            Texto = Replace(Texto, "    ", "")
            Texto = Replace(Texto, " -", vbLf & "- ")
            Texto = Replace(Texto, "que explotó", "que explote")
            Texto = Replace(Texto, "  ", vbLf)
            Texto = Replace(Texto, "??", "")
            Texto = Replace(Texto, "-" & vbLf, "-")
            
            
            Rango.Value = Texto
        End If
    Next Rango
    
     For Each Celda In Selection
        If Celda.HasFormula = False Then
            ' Verificar si la celda no contiene fórmulas
            Texto = Celda.Value
            
            ' Reemplazar el primer asterisco en cada línea por un guión
            Texto = Replace(Texto, vbLf & "*", vbLf & "-")
            
            ' Asignar el texto modificado de nuevo a la celda
            Celda.Value = Texto
        End If
    Next Celda
End Sub

