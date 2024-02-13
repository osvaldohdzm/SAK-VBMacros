Attribute VB_Name = "Módulo2"
Sub LimpiarReferencias()
    Dim rng As Range
    Dim cell As Range
    Dim urls() As String
    Dim i As Integer
    Dim newContent As String
    Dim url As String
    Dim uniqueUrls As Object
    Dim sortedUrls As Variant

    ' Selecciona el rango deseado
    Set rng = Selection

    ' Crea un objeto de diccionario para almacenar las URL únicas
    Set uniqueUrls = CreateObject("Scripting.Dictionary")

    ' Recorre cada celda en el rango
    For Each cell In rng
        ' Reiniciar el contenido de la celda
        newContent = ""

        ' Divide el contenido de la celda por el carácter de nueva línea (Chr(10))
        urls = Split(cell.Value, Chr(10))

        ' Recorre cada URL en el array y agrega al diccionario
        For i = LBound(urls) To UBound(urls)
            url = Trim(urls(i))
            If url <> "" And InStr(url, "wikipedia.org") = 0 And Not uniqueUrls.exists(url) Then
                uniqueUrls.Add url, Nothing
            End If
        Next i

        ' Ordena y convierte las claves del diccionario en un array
        sortedUrls = SortArray(uniqueUrls.keys)

        ' Construye el nuevo contenido de la celda sin duplicados y ordenado alfabéticamente
        For i = LBound(sortedUrls) To UBound(sortedUrls)
            newContent = newContent & sortedUrls(i)
            If i < UBound(sortedUrls) Then
                newContent = newContent & Chr(10)
            End If
        Next i

        ' Guarda el contenido filtrado en la celda
        cell.Value = newContent

        ' Limpia el diccionario para la próxima celda
        uniqueUrls.RemoveAll
    Next cell
End Sub

Function SortArray(arr As Variant) As Variant
    Dim temp As Variant
    Dim i As Integer, j As Integer
    
    ' Ordena el array alfabéticamente
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    
    ' Devuelve el array ordenado
    SortArray = arr
End Function

