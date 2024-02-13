Attribute VB_Name = "Módulo1"
Sub FiltrarURLsMejorado()
    ' Declaración de variables
    Dim rng As Range
    Dim cell As Range
    Dim urls() As String
    Dim i As Integer
    Dim newContent As String
    Dim url As String
    Dim delimiter As Variant
    Dim uniqueURLs
    Set uniqueURLs = CreateObject("Scripting.Dictionary") ' Objeto para almacenar URLs únicas


    ' Seleccionar el rango deseado
    Set rng = Selection

    ' Definir los delimitadores
    delimiter = Array(vbCrLf, vbLf, vbCr, vbCrLf & vbCrLf, vbCrLf & vbLf, vbLf & vbCrLf)

    ' Recorrer cada celda en el rango
    For Each cell In rng
        ' Reiniciar el contenido de la celda
        newContent = ""

        ' Dividir el contenido de la celda por cada delimitador
        For Each del In delimiter
            urls = Split(cell.Value, del, -1, vbTextCompare)

            For i = LBound(urls) To UBound(urls)
                url = Trim(urls(i))

                If url <> "" And Not uniqueURLs.Exists(url) Then
                    If InStr(url, "wikipedia.org") = 0 Then
                        uniqueURLs.Add url, Nothing
                        newContent = newContent & url & vbCrLf
                    End If
                End If
            Next i
        Next del

        ' Eliminar duplicados y espacios en blanco
        newContent = Trim(Join(uniqueURLs.Keys, vbCrLf))

        ' Guardar el contenido filtrado en la celda
        cell.Value = newContent

        ' Limpiar el diccionario para la siguiente celda
        uniqueURLs.RemoveAll
    Next cell
End Sub

