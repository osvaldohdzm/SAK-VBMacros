Attribute VB_Name = "M�dulo1"
Dim service_urls As Variant

Sub TraducirCeldasSeleccionadas()
    Dim celda As Range
    Dim textoOriginal As String
    Dim textoTraducido As String
    
    ' Establecer el idioma de origen y destino
    Dim idiomaOrigen As String
    Dim idiomaDestino As String
    idiomaOrigen = "en"
    idiomaDestino = "es"
    
    ' Definir la lista de servidores de traducci�n
    service_urls = Array( _
        "translate.google.com.mx", _
        "translate.google.fi", _
        "translate.google.fm", _
        "translate.google.fr", _
        "translate.google.com.co", _
        "translate.google.us", _
        "translate.google.ca", _
        "translate.google.es", _
        "translate.google.de" _
    )
    
    ' Definir el n�mero m�ximo de peticiones por grupo
    Dim maxRequestsPerGroup As Integer
    maxRequestsPerGroup = 30
    
    ' Inicializar contador para controlar el n�mero de peticiones en cada grupo
    Dim requestCount As Integer
    requestCount = 0
    
    ' Inicializar el �ndice para seleccionar un servidor de traducci�n de la lista
    Dim serverIndex As Integer
    serverIndex = 0
    
    ' Obtener el n�mero total de celdas seleccionadas
    Dim totalCeldas As Integer
    totalCeldas = Selection.Count
    
    ' Imprimir informaci�n en el Inmediato
    Debug.Print "N�mero total de celdas seleccionadas: " & totalCeldas
    
    ' Recorrer todas las celdas seleccionadas en la hoja activa
    For Each celda In Selection
        ' Obtener el texto original de la celda
        textoOriginal = celda.Value
        
        ' Verificar si la celda no est� vac�a
        If textoOriginal <> "" Then
            ' Almacenar el resultado de EncodeURL en una variable
            Dim textoCodificado As String
            textoCodificado = WorksheetFunction.EncodeURL(textoOriginal)
            
            ' Traducir el texto utilizando la funci�n translate_text
            textoTraducido = translate_text(textoCodificado, idiomaOrigen, idiomaDestino, service_urls(serverIndex))
            
            ' Colocar el texto traducido en la misma celda
            celda.Value = textoTraducido
            
            ' Incrementar el contador de peticiones en el grupo
            requestCount = requestCount + 1
            
            ' Imprimir informaci�n en el Inmediato
            Debug.Print "Celda traducida: " & celda.Address & " - Texto traducido: " & textoTraducido
            
            ' Verificar si se alcanz� el l�mite de peticiones por grupo
            If requestCount = maxRequestsPerGroup Then
                ' Reiniciar el contador y pasar al siguiente servidor
                requestCount = 0
                serverIndex = (serverIndex + 1) Mod UBound(service_urls) + 1
            End If
        End If
    Next celda
End Sub

Function translate_text(text_str As String, src_lang As String, trgt_lang As String, ByVal service_url As String) As String
    Dim url_str As String
    Dim xmlhttp As Object
    Dim responseText As String
    Const url_temp_src As String = "https://translate.googleapis.com/translate_a/single?client=gtx&sl=[from]&tl=[to]&dt=t&q="
    
    ' Construir la URL con el servicio espec�fico
    url_str = url_temp_src & text_str
    url_str = Replace(url_str, "[to]", trgt_lang)
    url_str = Replace(url_str, "[from]", src_lang)
    
    ' Crear un objeto XMLHTTP
    Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' Realizar la solicitud HTTP
    xmlhttp.Open "GET", url_str, False
    xmlhttp.send
    
    ' Obtener la respuesta
    responseText = xmlhttp.responseText
    
    ' Traducir la respuesta utilizando ParseTranslationResponse
    translate_text = ParseTranslationResponse(responseText)
End Function

Function ParseTranslationResponse(responseText As String) As String
    ' Analizar la respuesta JSON para obtener el texto traducido
    Dim startIndex As Long
    Dim endIndex As Long
    
    ' Buscar el inicio de la cadena de traducci�n
    startIndex = InStr(responseText, "[""") + 2
    
    ' Buscar el final de la cadena de traducci�n
    endIndex = InStr(startIndex, responseText, """,")
    
    ' Extraer la cadena de traducci�n
    If startIndex > 0 And endIndex > startIndex Then
        ' Obtener la cadena de traducci�n con caracteres especiales reemplazados
        Dim translatedText As String
        translatedText = Mid(responseText, startIndex, endIndex - startIndex)
        translatedText = Replace(translatedText, "\u003c", "<")
        translatedText = Replace(translatedText, "\u003e", ">")
        
        ParseTranslationResponse = translatedText
    Else
        ParseTranslationResponse = "Error al analizar la respuesta"
    End If
End Function


