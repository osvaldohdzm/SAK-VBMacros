Attribute VB_Name = "M�dulo1"
Sub AjustarAlturaConMargen()
    Dim celda As Range
    Dim alturaAntigua As Double
    
    ' Desactivar el ajuste autom�tico de altura
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Recorrer cada celda en la selecci�n
    For Each celda In Selection
        ' Guardar la altura original de la fila
        alturaAntigua = celda.RowHeight
        
        ' Ajustar la altura de la fila al contenido
        celda.EntireRow.AutoFit
        
        ' Agregar 3 puntos de altura adicional
        celda.RowHeight = celda.RowHeight + 3
    Next celda
    
    ' Reactivar ajuste autom�tico de altura
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

