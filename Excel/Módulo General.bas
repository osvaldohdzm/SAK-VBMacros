Attribute VB_Name = "ExcelMacrosGeneral"



Sub Lowercase()
 For Each cell In Selection
        If Not cell.HasFormula Then
            cell.Value = LCase(cell.Value)
        End If
    Next cell
End Sub

Sub AjustarAlturaFilasEnTodasLasHojasDelLibroActivo()
    Dim sh As Worksheet
    
    ' Recorre todas las hojas en el libro activo
    For Each sh In ActiveWorkbook.Worksheets
        ' Ajusta la altura de todas las filas en la hoja actual
        sh.Rows.RowHeight = 15
    Next sh
    
    ' Muestra un mensaje indicando que la operaci�n se complet�
    MsgBox "Todas las filas en todas las hojas del libro activo se han ajustado a una altura de 15."
End Sub


Sub EliminarSaltosDeLinea()

    Dim celda As Range
    Dim Texto As String
    Dim NuevoTexto As String
    
    ' Itera a trav�s de las celdas seleccionadas en la hoja activa
    For Each celda In Selection
        If Not celda.HasFormula Then ' Ignora celdas con f�rmulas
            Texto = celda.Value
            
            ' Reemplazar diferentes tipos de saltos de l�nea y retornos de carro
            NuevoTexto = Replace(Texto, vbCrLf, " ")   ' Salto de l�nea + retorno de carro
            NuevoTexto = Replace(NuevoTexto, vbCr, " ") ' Retorno de carro
            NuevoTexto = Replace(NuevoTexto, vbLf, " ") ' Salto de l�nea
            
            celda.Value = NuevoTexto ' Asigna el nuevo valor a la celda
        End If
    Next celda

End Sub

Attribute VB_Name = "ExcelModuloGeneral"
Sub EliminarLineasVaciasEnCeldasSeleccionadas()
    Dim celda As Range
    Dim lineas As Variant
    Dim i As Integer

    ' Iterar sobre cada celda seleccionada
    For Each celda In Selection
        ' Verificar si la celda tiene texto
        If Not IsEmpty(celda.Value) Then
            ' Reemplazar diferentes saltos de l�nea con vbLf
            Dim contenido As String
            contenido = Replace(Replace(Replace(celda.Value, vbCrLf, vbLf), vbCr, vbLf), vbLf & vbLf, vbLf)
            
            ' Si el contenido comienza con vbLf, quitarlo
            If Left(contenido, 1) = vbLf Then
                contenido = Mid(contenido, 2)
            End If
            
            ' Si el contenido termina con vbLf, quitarlo
            If Right(contenido, 1) = vbLf Then
                contenido = Left(contenido, Len(contenido) - 1)
            End If
            
            ' Dividir el contenido de la celda en un array de l�neas
            lineas = Split(contenido, vbLf)
            
            ' Crear un nuevo array para almacenar las l�neas no vac�as
            Dim lineasSinVacias() As String
            ReDim lineasSinVacias(0 To UBound(lineas))
            Dim idx As Integer
            idx = 0
            
            ' Iterar sobre cada l�nea del array
            For i = LBound(lineas) To UBound(lineas)
                ' Verificar si la l�nea est� vac�a y no agregarla al nuevo array
                If Trim(lineas(i)) <> "" Then
                    lineasSinVacias(idx) = lineas(i)
                    idx = idx + 1
                End If
            Next i
            
            ' Redimensionar el array resultante
            ReDim Preserve lineasSinVacias(0 To idx - 1)
            
            ' Unir el array de l�neas de nuevo en una cadena y asignarlo a la celda
            celda.Value = Join(lineasSinVacias, vbLf)
        End If
    Next celda
End Sub

Sub ExportarTabla()
    Dim celdaActual As Range
    Dim tabla As ListObject
    Dim rutaArchivo As String
    Dim nombreArchivo As String
    Dim carpetaDestino As String
    Dim nuevoLibro As Workbook
    Dim nuevaHoja As Worksheet
    Dim archivoGuardado As Variant
    
    ' Obtener la celda actualmente seleccionada
    Set celdaActual = ActiveCell
    
    ' Verificar si la celda seleccionada est� dentro de una tabla
    On Error Resume Next
    Set tabla = celdaActual.ListObject
    On Error GoTo 0
    
    ' Si la celda est� dentro de una tabla, procedemos
    If Not tabla Is Nothing Then
        ' Obtener el nombre de la tabla
        nombreArchivo = tabla.Name
        
        ' Mostrar un cuadro de di�logo para seleccionar la carpeta de destino
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Selecciona la carpeta para guardar el archivo"
            .Show
            If .SelectedItems.Count > 0 Then
                carpetaDestino = .SelectedItems(1)
            Else
                MsgBox "No se seleccion� ninguna carpeta. La exportaci�n se ha cancelado.", vbExclamation
                Exit Sub
            End If
        End With
        
        ' Definir la ruta del archivo
        rutaArchivo = carpetaDestino & "\" & nombreArchivo & ".csv"
        
        ' Crear una nueva instancia de Excel
        Set nuevoLibro = Workbooks.Add
        Set nuevaHoja = nuevoLibro.Sheets(1)
        
        ' Copiar la tabla a la nueva hoja
        tabla.Range.Copy
        nuevaHoja.Cells.PasteSpecial xlPasteValues
        
        ' Guardar la nueva hoja como archivo CSV
        Application.DisplayAlerts = False
        nuevoLibro.SaveAs Filename:=rutaArchivo, FileFormat:=xlCSV, CreateBackup:=False
        Application.DisplayAlerts = True
        
        ' Cerrar la nueva instancia de Excel sin guardar cambios
        nuevoLibro.Close SaveChanges:=False
        
        MsgBox "Archivo exportado con �xito: " & rutaArchivo
        
        ' Regresar a la hoja original
        ThisWorkbook.Sheets(1).Activate
        
    Else
        MsgBox "La celda seleccionada no est� dentro de una tabla."
    End If
End Sub






Sub LimpiarEtiquetasHTML()
    Dim selectedRange As Range
    Dim cell As Range
    Dim htmlPattern As String
    Dim additionalPattern As String
    
    ' Definir el patr�n HTML que se desea eliminar
    htmlPattern = "<(\/?(p|a|li|ul|b|strong|i|u|br)[^>]*?)>|<\/p><p>"
    
    ' Definir el patr�n para eliminar etiquetas <div>, </div> y <span>, </span> pero mantener su contenido
    additionalPattern = "<(div|span)[^>]*>|<\/(div|span)>"
    
    ' Obtener el rango de celdas seleccionadas por el usuario
    On Error Resume Next
    Set selectedRange = Application.InputBox("Seleccione el rango de celdas", Type:=8)
    On Error GoTo 0
    
    ' Salir si el usuario cancela la selecci�n
    If selectedRange Is Nothing Then Exit Sub
    
    ' Iterar sobre cada celda en el rango seleccionado
    For Each cell In selectedRange
        ' Verificar si la celda contiene texto
        If Not IsEmpty(cell.Value) And TypeName(cell.Value) = "String" Then
            ' Eliminar las etiquetas HTML utilizando expresiones regulares
            cell.Value = RegExpReplace(cell.Value, htmlPattern, vbCrLf) ' Reemplazar con salto de l�nea
            ' Adem�s, eliminar las etiquetas <div>, </div>, <span> y </span> pero mantener su contenido
            cell.Value = RegExpReplace(cell.Value, additionalPattern, "")
        End If
    Next cell
    
    MsgBox "Etiquetas HTML eliminadas correctamente y reemplazadas seg�n lo solicitado.", vbInformation
End Sub

Function RegExpReplace(ByVal text As String, ByVal replacePattern As String, ByVal replaceWith As String) As String
    ' Funci�n para reemplazar utilizando expresiones regulares
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = replacePattern
    End With
    
    RegExpReplace = regex.Replace(text, replaceWith)
End Function




Sub TraducirCeldasSeleccionadas()
    Dim celda As Range
    Dim textoOriginal As String
    Dim textoTraducido As String
    Dim service_urls As Variant
    
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
    Dim spanishText As String
    Dim posStart As Long
    Dim posEnd As Long
    Dim tempText As String
    Dim isHash As Boolean

    ' Inicializar la variable para acumular el texto en espa�ol
    spanishText = ""

    ' Inicializar la posici�n de b�squeda
    posStart = 1

    ' Buscar y extraer el texto en espa�ol
    Do
        ' Buscar el inicio de la cadena de texto en espa�ol
        posStart = InStr(posStart, responseText, "[""")
        If posStart = 0 Then Exit Do
        posStart = posStart + 2

        ' Buscar el final de la cadena de texto en espa�ol
        posEnd = InStr(posStart, responseText, """,")
        If posEnd = 0 Then Exit Do

        ' Extraer el texto en espa�ol
        tempText = Mid(responseText, posStart, posEnd - posStart)
        tempText = Replace(tempText, "\", "") ' Limpiar caracteres de escape
        
        ' Verificar si el texto es un hash
        isHash = CheckIfHash(tempText)
        
        ' Si el texto no es un hash, a�adirlo al texto en espa�ol
        If Not isHash Then
            spanishText = spanishText & tempText & " "
        End If

        ' Mover la posici�n de b�squeda para el pr�ximo par
        posStart = posEnd + 2
    Loop

    ' Eliminar el �ltimo espacio en blanco a�adido
    If Len(spanishText) > 0 Then
        spanishText = Trim(spanishText)
    End If

    ' Retornar el texto en espa�ol extra�do
    ParseTranslationResponse = spanishText
End Function

Function CheckIfHash(text As String) As Boolean
    ' Verificar si el texto parece un hash MD5 (32 caracteres hexadecimales)
    Dim pattern As String
    Dim regex As Object
    
    pattern = "^[a-fA-F0-9]{32}$" ' Patr�n para un hash MD5
    
    ' Crear objeto de expresi�n regular
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = False
        .IgnoreCase = True
        .pattern = pattern
    End With
    
    ' Devolver True si el texto coincide con el patr�n de hash
    CheckIfHash = regex.Test(text)
End Function







