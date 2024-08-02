Attribute VB_Name = "Módulo1"
Sub ConvertSelectionToTextFormat()
    Dim cell As Range
    
    ' Desactivar la actualización de la pantalla para mejorar la eficiencia
    Application.ScreenUpdating = False
    
    ' Iterar a través de cada celda en la selección
    For Each cell In Selection
        If IsDate(cell.Value) Then
            ' Convertir la fecha a texto en el formato deseado y almacenar en una variable
            Dim formattedDate As String
            formattedDate = Format(cell.Value, "yyyy-mm-dd")
            
            ' Establecer el formato de la celda a texto
            cell.NumberFormat = "@"
            
            ' Asignar el valor formateado a la celda
            cell.Value = formattedDate
        End If
    Next cell
    
    ' Reactivar la actualización de la pantalla
    Application.ScreenUpdating = True
    
    MsgBox "Conversión a formato de texto completada."
End Sub

