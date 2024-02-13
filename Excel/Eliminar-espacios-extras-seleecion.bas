
Sub QuitarEspacios()
    Dim rng As Range
    Dim c As Range
    
    Set rng = Selection 'asume que el rango seleccionado es el que quieres modificar
    
    For Each c In rng 'recorre cada celda del rango
        c.Value = Application.Trim(c.Value) 'quita los espacios de la celda
    Next c
End Sub
