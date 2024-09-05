Attribute VB_Name = "ExcelMacrosCibersecurity"
Sub ReemplazarPalabras()
    Dim c As Range
    Dim valorActual As String
    
    For Each c In Selection
        valorActual = Trim(UCase(c.Value)) ' Convertimos a mayúsculas y eliminamos espacios adicionales
        
        Select Case valorActual
            Case "0", "NONE", "INFORMATIVA", "INFO"
                c.Value = "INFORMATIVA"
            Case "1", "BAJA", "BAJO", "LOW"
                c.Value = "BAJA"
            Case "2", "BAJA", "BAJO", "LOW"
                c.Value = "BAJA"
            Case "3", "BAJA", "BAJO", "LOW"
                c.Value = "BAJA"
            Case "4", "BAJA", "BAJO", "LOW"
                c.Value = "BAJA"
            Case "5", "MEDIA", "MEDIO", "MEDIUM"
                c.Value = "MEDIA"
            Case "6", "MEDIA", "MEDIO", "MEDIUM"
                c.Value = "MEDIA"
            Case "7", "ALTO", "HIGH"
                c.Value = "ALTA"
            Case "8", "ALTA", "HIGH"
                c.Value = "ALTA"
            Case "9", "CRÍTICA", "CRITICAL", "CRÍTICO"
                c.Value = "CRÍTICA"
            Case "10", "CRÍTICA", "CRITICAL", "CRÍTICO"
                c.Value = "CRÍTICA"
            ' Mantener el contenido actual si no coincide con las palabras a reemplazar
            Case Else
                ' No hacer nada
        End Select
    Next c
End Sub



Sub LimpiarCeldasYMostrarContenidoComoArray()
    Dim rng As Range
    Dim cell As Range
    Dim content As String
    Dim contentArray() As String
    Dim i As Integer
    Dim temp As String
    Dim uniqueUrls As Object
    Dim uniqueArray() As String
    Dim n As Integer
    
    ' Selecciona el rango deseado
    Set rng = Selection
    
    ' Recorre cada celda en el rango
    For Each cell In rng
        ' Obtiene el contenido de la celda
        content = cell.Value
        
        ' Comprueba si el contenido es vacío
        If content <> "" Then
            ' Convierte el contenido en un array separado por el carácter de nueva línea
            contentArray = Split(content, Chr(10))
            
            ' Inicializa el diccionario para almacenar las URL únicas
            Set uniqueUrls = CreateObject("Scripting.Dictionary")
            
            ' Agrega las URL únicas al diccionario
            For i = LBound(contentArray) To UBound(contentArray)
                If contentArray(i) <> "" Then
                    ' Elimina espacios en blanco, Chr(10) y Chr(13) del elemento
                    contentArray(i) = Trim(Replace(contentArray(i), Chr(10), ""))
                    contentArray(i) = Trim(Replace(contentArray(i), Chr(13), ""))
                    contentArray(i) = Replace(contentArray(i), " ", "")
                    If InStr(1, contentArray(i), "wikipedia", vbTextCompare) = 0 Then
                        If Not uniqueUrls.exists(contentArray(i)) Then
                            uniqueUrls.Add contentArray(i), Nothing
                        End If
                    End If
                End If
            Next i
            
            ' Convertir la colección de claves en un array
            n = uniqueUrls.Count - 1
            ReDim uniqueArray(n)
            i = 0
            For Each Key In uniqueUrls.Keys
                uniqueArray(i) = Key
                i = i + 1
            Next
            
            ' Ordena los elementos del array
            For i = LBound(uniqueArray) To UBound(uniqueArray) - 1
                For j = i + 1 To UBound(uniqueArray)
                    If uniqueArray(i) > uniqueArray(j) Then
                        temp = uniqueArray(i)
                        uniqueArray(i) = uniqueArray(j)
                        uniqueArray(j) = temp
                    End If
                Next j
            Next i
            
            ' Convierte el array nuevamente en una cadena concatenada por el carácter de nueva línea
            content = Join(uniqueArray, Chr(10))
            
            ' Asigna el contenido filtrado a la celda
            cell.Value = content
        End If
    Next cell
End Sub

