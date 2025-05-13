Attribute VB_Name = "Módulo1"
Sub EliminarDuplicados()
    Dim rng As Range
    Dim cell As Range
    Dim urls() As String
    Dim i As Integer
    Dim uniqueUrls As Object
    Dim url As String
    Dim delimiter As Variant

    ' Selecciona el rango deseado
    Set rng = Selection

    ' Define los delimitadores para dividir el contenido de las celdas
    delimiter = Array(vbCrLf, vbLf, vbCr, vbCrLf & vbCrLf, vbCrLf & vbLf, vbLf & vbCrLf)

    ' Crear un objeto de diccionario para almacenar las URL únicas
    Set uniqueUrls = CreateObject("Scripting.Dictionary")

    ' Recorre cada celda en el rango
    For Each cell In rng
        ' Divide el contenido de la celda por cada delimitador
        For Each del In delimiter
            urls = Split(cell.Value, del, -1, vbTextCompare)
            
            ' Recorre cada URL en el array
            For i = LBound(urls) To UBound(urls)
                url = Trim(urls(i))
                ' Verifica si la URL no está en el diccionario de URL únicas
                If Not uniqueUrls.exists(url) Then
                    ' Si la URL no está en el diccionario, la agrega
                    uniqueUrls.Add url, Nothing
                End If
            Next i
        Next del
        
        ' Restablece el contenido de la celda
        cell.Value = Join(uniqueUrls.keys, vbCrLf)
        
        ' Restablece el diccionario para la próxima celda
        uniqueUrls.RemoveAll
    Next cell
End Sub

