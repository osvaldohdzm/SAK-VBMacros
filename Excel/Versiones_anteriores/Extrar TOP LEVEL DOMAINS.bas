Attribute VB_Name = "Módulo1"
Sub MantenerDominioSinSubdominio()
    Dim cell As Range
    Dim dominio As String
    Dim partes() As String
    Dim resultado As String
    
    ' Recorrer cada celda seleccionada
    For Each cell In Selection
        ' Obtener el contenido de la celda
        dominio = cell.Value
        
        ' Dividir el dominio en partes (subdominio y dominio principal)
        partes = Split(dominio, ".")
        
        ' Construir el resultado con los tres últimos elementos del dominio
        resultado = partes(UBound(partes) - 2) & "." & partes(UBound(partes) - 1) & "." & partes(UBound(partes))
        
        ' Reemplazar el contenido de la celda con el resultado
        cell.Value = resultado
    Next cell
    
    MsgBox "Se ha eliminado el subdominio de las celdas seleccionadas.", vbInformation
End Sub

