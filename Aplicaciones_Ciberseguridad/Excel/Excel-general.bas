Attribute VB_Name = "M�dulo1"
Sub ConvertSelectionToTextFormat()
    Dim cell As Range
    
    ' Desactivar la actualizaci�n de la pantalla para mejorar la eficiencia
    Application.ScreenUpdating = False
    
    ' Iterar a trav�s de cada celda en la selecci�n
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
    
    ' Reactivar la actualizaci�n de la pantalla
    Application.ScreenUpdating = True
    
    MsgBox "Conversi�n a formato de texto completada."
End Sub

