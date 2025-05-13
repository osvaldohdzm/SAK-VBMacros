Attribute VB_Name = "M�dulo1"
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

    ' Crear un objeto de diccionario para almacenar las URL �nicas
    Set uniqueUrls = CreateObject("Scripting.Dictionary")

    ' Recorre cada celda en el rango
    For Each cell In rng
        ' Divide el contenido de la celda por cada delimitador
        For Each del In delimiter
            urls = Split(cell.Value, del, -1, vbTextCompare)
            
            ' Recorre cada URL en el array
            For i = LBound(urls) To UBound(urls)
                url = Trim(urls(i))
                ' Verifica si la URL no est� en el diccionario de URL �nicas
                If Not uniqueUrls.exists(url) Then
                    ' Si la URL no est� en el diccionario, la agrega
                    uniqueUrls.Add url, Nothing
                End If
            Next i
        Next del
        
        ' Restablece el contenido de la celda
        cell.Value = Join(uniqueUrls.keys, vbCrLf)
        
        ' Restablece el diccionario para la pr�xima celda
        uniqueUrls.RemoveAll
    Next cell
End Sub

