Attribute VB_Name = "M�dulo1"
Sub MantenerDominioSinSubdominio()
    Dim cell As Range
    Dim dominio As String
    Dim partes() As String
    Dim resultado As String
    
    ' Recorrer cada celda seleccionada
    For Each cell In Selection
        ' Obtener el contenido de la celda
        dominio = cell.Value
        
        ' Dividir el dominio en partes (subdominios separados por punto)
        partes = Split(dominio, ".")
        
        ' Verificar si hay suficientes partes para extraer el dominio principal y el TLD
        If UBound(partes) >= 2 Then
            ' Determinar el �ndice desde el cual comienza el dominio principal
            Dim startIndex As Integer
            startIndex = UBound(partes) - 2 ' Comienza desde el tercer �ltimo elemento
            
            ' Construir el resultado usando el dominio principal y el TLD (top-level domain)
            resultado = partes(startIndex) & "." & partes(startIndex + 1) & "." & partes(startIndex + 2)
        Else
            ' Si no hay suficientes partes, mantener el dominio original
            resultado = dominio
        End If
        
        ' Reemplazar el contenido de la celda con el resultado
        cell.Value = resultado
    Next cell
    
    MsgBox "Se ha extra�do el dominio principal de las celdas seleccionadas.", vbInformation
End Sub

