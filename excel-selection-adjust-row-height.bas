Attribute VB_Name = "Módulo1"
Sub AjustarAlturaConMargen()
    Dim celda As Range
    Dim alturaAntigua As Double
    
    ' Desactivar el ajuste automático de altura
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Recorrer cada celda en la selección
    For Each celda In Selection
        ' Guardar la altura original de la fila
        alturaAntigua = celda.RowHeight
        
        ' Ajustar la altura de la fila al contenido
        celda.EntireRow.AutoFit
        
        ' Agregar 3 puntos de altura adicional
        celda.RowHeight = celda.RowHeight + 3
    Next celda
    
    ' Reactivar ajuste automático de altura
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

