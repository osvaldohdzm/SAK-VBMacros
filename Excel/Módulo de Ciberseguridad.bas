Attribute VB_Name = "ExcelMacrosCibersecurity"

Sub ReemplazarCadenasSeveridades()

    Dim c As Range
    Dim valorActual As String
    
    For Each c In Selection
        valorActual = Trim(UCase(c.value)) ' Convertimos a mayúsculas y eliminamos espacios adicionales
        
        Select Case valorActual
            Case "0", "NONE", "INFORMATIVA", "INFO"
                c.value = "INFORMATIVA"
            Case "1", "BAJA", "BAJO", "LOW"
                c.value = "BAJA"
            Case "2", "BAJA", "BAJO", "LOW"
                c.value = "BAJA"
            Case "3", "BAJA", "BAJO", "LOW"
                c.value = "BAJA"
            Case "4", "BAJA", "BAJO", "LOW"
                c.value = "BAJA"
            Case "5", "MEDIA", "MEDIO", "MEDIUM"
                c.value = "MEDIA"
            Case "6", "MEDIA", "MEDIO", "MEDIUM"
                c.value = "MEDIA"
            Case "7", "ALTO", "HIGH"
                c.value = "ALTA"
            Case "8", "ALTA", "HIGH"
                c.value = "ALTA"
            Case "9", "CRÍTICA", "CRITICAL", "CRÍTICO"
                c.value = "CRÍTICA"
            Case "10", "CRÍTICA", "CRITICAL", "CRÍTICO"
                c.value = "CRÍTICA"
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
        content = cell.value
        
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
                        If Not uniqueUrls.Exists(contentArray(i)) Then
                            uniqueUrls.Add contentArray(i), Nothing
                        End If
                    End If
                End If
            Next i
            
            ' Convertir la colección de claves en un array
            n = uniqueUrls.Count - 1
            ReDim uniqueArray(n)
            i = 0
            For Each key In uniqueUrls.keys
                uniqueArray(i) = key
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
            cell.value = content
        End If
    Next cell
End Sub

Sub ReplaceWithURLs()
    Dim cell As Range
    Dim parts As Variant
    Dim url As String
    Dim i As Integer
    
    ' Recorre cada celda en el rango seleccionado
    For Each cell In Selection
        If cell.value <> "" Then
            ' Separa la cadena por comas
            parts = Split(cell.value, ",")
            
            ' Inicializa una cadena vacía para las URLs
            url = ""
            
            ' Recorre cada parte de la cadena
            For i = LBound(parts) To UBound(parts)
                ' Separa cada parte por el símbolo |
                If InStr(parts(i), "|") > 0 Then
                    url = url & Mid(parts(i), InStr(parts(i), "|") + 1) & vbLf
                End If
            Next i
            
            ' Elimina el último salto de línea
            If Len(url) > 0 Then
                url = Left(url, Len(url) - 1)
            End If
            
            ' Reemplaza las comillas dobles sobrantes
            url = Replace(url, """", "")
            
            ' Reemplaza el valor de la celda con las URLs y saltos de línea
            cell.value = url
        End If
    Next cell
End Sub

Sub AplicarFormatoCondicional()
    Dim selectedRange As Range

    ' Verificar si hay celdas seleccionadas
    On Error Resume Next
    Set selectedRange = Selection.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0

    ' Salir si no hay celdas seleccionadas
    If selectedRange Is Nothing Then
        MsgBox "No hay celdas seleccionadas."
        Exit Sub
    End If

    ' Aplicar formato condicional según el contenido de las celdas seleccionadas
    With selectedRange
        .FormatConditions.Add Type:=xlTextString, String:="CRÍTICA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(255, 255, 255) ' Blanco
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(112, 48, 160) ' #7030A0
        End With

        .FormatConditions.Add Type:=xlTextString, String:="ALTA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(255, 255, 255) ' Blanco
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(255, 0, 0) ' #FF0000
        End With

        .FormatConditions.Add Type:=xlTextString, String:="MEDIA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(0, 0, 0) ' Negro
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(255, 255, 0) ' #FFFF00
        End With

        .FormatConditions.Add Type:=xlTextString, String:="BAJA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(255, 255, 255) ' Blanco
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(0, 176, 80) ' #00B050
        End With

        .FormatConditions.Add Type:=xlTextString, String:="INFORMATIVA", TextOperator:=xlContains
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Font
            .Color = RGB(0, 0, 0) ' Negro
        End With
        With .FormatConditions(1).Interior
            .Color = RGB(231, 230, 230) ' #E7E6E6
        End With
    End With
End Sub



Sub ConvertirATextoEnOracion()
    Dim celda As Range
    Dim Texto As String
    Dim primeraLetra As String
    Dim restoTexto As String

    ' Recorre cada celda en el rango seleccionado
    For Each celda In Selection
        If Not IsEmpty(celda.value) Then
            Texto = celda.value
            ' Convierte todo el texto a minúsculas
            Texto = LCase(Texto)
            ' Extrae la primera letra
            primeraLetra = UCase(Left(Texto, 1))
            ' Extrae el resto del texto
            restoTexto = Mid(Texto, 2)
            ' Combina la primera letra en mayúsculas con el resto del texto en minúsculas
            celda.value = primeraLetra & restoTexto
        End If
    Next celda
End Sub


Sub QuitarEspacios()
    Dim rng As Range
    Dim c As Range
    
    Set rng = Selection 'asume que el rango seleccionado es el que quieres modificar
    
    For Each c In rng 'recorre cada celda del rango
        c.value = Application.Trim(c.value) 'quita los espacios de la celda
    Next c
End Sub

Sub LimpiarSalida()
    Dim rng As Range
    Dim celda As Range
    Dim lineas As Variant
    Dim i As Integer
    Dim htmlPattern As String
    Dim liPattern As String
    Dim cleanHtmlPattern As String
    
    ' Definir los patrones HTML
    htmlPattern = "<(\/?(p|a|ul|b|strong|i|u|br)[^>]*?)>|<\/p><p>"
    liPattern = "<li[^>]*?>"
    cleanHtmlPattern = "<[^>]+>" ' Para eliminar cualquier etiqueta HTML y sus atributos
    
    ' Asume que el rango seleccionado es el que quieres modificar
    Set rng = Selection
    
    ' Reemplaza caracteres de tabulación con espacios y quita espacios en blanco adicionales
    For Each celda In rng
        If Not IsEmpty(celda.value) Then
            ' Reemplazar caracteres de tabulación con espacios
            celda.value = Replace(celda.value, Chr(9), " ")
            ' Quitar espacios en blanco adicionales
            celda.value = Application.Trim(celda.value)
            
                        ' Reemplazar <li> etiquetas con saltos de línea
            celda.value = RegExpReplace(celda.value, liPattern, vbLf)
            
            ' Eliminar otras etiquetas HTML pero conservar el texto interno
            celda.value = RegExpReplace(celda.value, cleanHtmlPattern, vbNullString)
            
            ' Unificar saltos de línea y eliminar líneas vacías
            lineas = Split(celda.value, vbLf)
            For i = LBound(lineas) To UBound(lineas)
                If Trim(lineas(i)) = "" Then
                    lineas(i) = vbNullString
                End If
            Next i
            
            ' Unir líneas no vacías y eliminar posibles saltos de línea finales
            celda.value = Join(lineas, vbLf)
            If Right(celda.value, 1) = vbLf Then
                celda.value = Left(celda.value, Len(celda.value) - 1)
            End If
            
            ' Eliminar etiquetas HTML que no se hayan eliminado y reemplazar con saltos de línea
            celda.value = RegExpReplace(celda.value, htmlPattern, vbLf)
            
            ' Reemplazar entidades HTML con caracteres correspondientes
            celda.value = ReplaceHtmlEntities(celda.value)
            
        End If
    Next celda
    
    MsgBox "Proceso completado: Espacios, etiquetas HTML y saltos de línea ajustados.", vbInformation
End Sub

Function RegExpReplace(ByVal text As String, ByVal replacePattern As String, ByVal replaceWith As String) As String
    ' Función para reemplazar utilizando expresiones regulares
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = replacePattern
    End With
    
    RegExpReplace = regex.Replace(text, replaceWith)
End Function

Function ReplaceHtmlEntities(ByVal text As String) As String
    ' Función para reemplazar entidades HTML con caracteres correspondientes
    text = Replace(text, "&lt;", "<")
    text = Replace(text, "&gt;", ">")
    text = Replace(text, "&amp;", "&")
    text = Replace(text, "&quot;", """")
    text = Replace(text, "&apos;", "'")
    text = Replace(text, "&#x27;", "'")
    text = Replace(text, "&#34;", """")
    text = Replace(text, "&#39;", "'")
    text = Replace(text, "&#160;", Chr(160)) ' Espacio no separable
    
    ReplaceHtmlEntities = text
End Function



Sub OrdenaSegunColorRelleno()
    Dim celdaActual As Range
    Dim tabla As ListObject
    Dim ws As Worksheet
    Dim respuesta As VbMsgBoxResult

    ' Obtener la celda actualmente seleccionada
    Set celdaActual = ActiveCell
    
    ' Obtener la hoja activa
    Set ws = ActiveSheet
    
    ' Verificar si la celda seleccionada está dentro de una tabla
    On Error Resume Next
    Set tabla = celdaActual.ListObject
    On Error GoTo 0
    
    ' Confirmar si la tabla encontrada es la tabla de vulnerabilidades
    If Not tabla Is Nothing Then
        ' Mostrar mensaje de confirmación
        respuesta = MsgBox("¿Estás seguro de que estás en una tabla de vulnerabilidades? Procederá a ordenar por color de relleno en la columna 'Severidad'.", vbYesNo + vbQuestion, "Confirmación")
        
        ' Si el usuario elige 'Sí', proceder con la ordenación
        If respuesta = vbYes Then
            With ws.ListObjects(tabla.Name).Sort
                ' Limpiar campos de ordenación previos
                .SortFields.Clear
                
                ' Ordenar por color de relleno verde
                .SortFields.Add key:=ws.ListObjects(tabla.Name).ListColumns("Severidad").Range, _
                    SortOn:=xlSortOnCellColor, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal
                .SortFields(1).SortOnValue.Color = RGB(0, 176, 80)
                .Apply
                
                ' Limpiar campos de ordenación previos
                .SortFields.Clear
                
                ' Ordenar por color de relleno amarillo
                .SortFields.Add key:=ws.ListObjects(tabla.Name).ListColumns("Severidad").Range, _
                    SortOn:=xlSortOnCellColor, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal
                .SortFields(1).SortOnValue.Color = RGB(255, 255, 0)
                .Apply
                
                ' Limpiar campos de ordenación previos
                .SortFields.Clear
                
                ' Ordenar por color de relleno rojo
                .SortFields.Add key:=ws.ListObjects(tabla.Name).ListColumns("Severidad").Range, _
                    SortOn:=xlSortOnCellColor, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal
                .SortFields(1).SortOnValue.Color = RGB(255, 0, 0)
                .Apply
            End With
        End If
    Else
        MsgBox "La celda seleccionada no está dentro de una tabla.", vbExclamation, "Error"
    End If
End Sub
' GenerarDocumentosWord

Sub WordAppReplaceParagraph(WordApp As Object, WordDoc As Object, wordToFind As String, replaceWord As String)
    Dim findInRange As Boolean
    
      ' Ir al principio del documento nuevamente
    WordApp.Selection.GoTo What:=1, Which:=1, Name:="1"
    WordApp.ActiveWindow.ActivePane.View.SeekView = 0
    
    ' Bucle para buscar y reemplazar todas las ocurrencias
    Do
        ' Intentar encontrar y reemplazar en el cuerpo del documento
        findInRange = WordApp.Selection.Find.Execute(findText:=wordToFind)
        
        ' Si se encontró el texto, reemplazarlo
        If findInRange Then
        
       
    
            ' Realizar el reemplazo
            WordApp.Selection.text = replaceWord
            
             ' Ir al principio del documento nuevamente
    WordApp.Selection.GoTo What:=1, Which:=1, Name:="1"
        End If
    Loop While findInRange
    
    
End Sub

Sub FormatRiskLevelCell(cell As Object)
    Dim cellText As String
    ' Obtener el texto de la celda y eliminar los caracteres especiales
    cellText = Replace(cell.Range.text, vbCrLf, "")
    cellText = Replace(cellText, vbCr, "")
    cellText = Replace(cellText, vbLf, "")
    cellText = Replace(cellText, Chr(7), "")
    
    ' Realizar la comparación utilizando el texto de la celda sin caracteres especiales
    Select Case cellText
        Case "CRÍTICA"
            cell.Shading.BackgroundPatternColor = 10498160
            cell.Range.Font.Color = 16777215
        Case "ALTA"
            cell.Shading.BackgroundPatternColor = 255
            cell.Range.Font.Color = 16777215
        Case "MEDIA"
            cell.Shading.BackgroundPatternColor = 65535
            cell.Range.Font.Color = 0
        Case "BAJA"
            cell.Shading.BackgroundPatternColor = 5287936
            cell.Range.Font.Color = 16777215
    End Select
End Sub

Function TransformText(text As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Configurar la expresión regular para encontrar saltos de línea o saltos de carro sin un punto antes y no seguidos de paréntesis ni de guión
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = "([^.()\r\n-])(?![^(]*\)|[-])[^\S\r\n]*[\r\n]+" ' Expresión regular para encontrar saltos de línea o saltos de carro sin un punto antes y no seguidos de paréntesis ni de guión
    End With
    
    ' Realizar la transformación: quitar caracteres especiales y aplicar la expresión regular
    TransformText = regex.Replace(Replace(text, Chr(7), ""), "$1 ")
End Function
Sub EliminarUltimasFilasSiEsSalidaPruebaSeguridad(WordDoc As Object, replaceDic As Object)
    Dim salidaPruebaSeguridadKey As String
    salidaPruebaSeguridadKey = "«Salidas de herramienta»"
    
    ' Verificar si la clave está presente en el diccionario
    If replaceDic.Exists(salidaPruebaSeguridadKey) Then
        ' Convertir el valor asociado a una cadena
        Dim keyValue As String
        keyValue = CStr(replaceDic(salidaPruebaSeguridadKey))
   
        
        ' Verificar si el valor es vacío, una cadena vacía o Null
        If Len(Trim(keyValue)) = 0 Then
            ' Eliminar las últimas dos filas de la primera tabla en el documento
            Dim firstTable As Object
            Set firstTable = WordDoc.Tables(1)
            Dim numRows As Integer
            numRows = firstTable.Rows.Count
            
            If numRows > 0 Then
                ' Eliminar la última fila dos veces si hay suficientes filas
                If numRows > 1 Then
                    firstTable.Rows(numRows).Delete
                    firstTable.Rows(numRows - 1).Delete
                ElseIf numRows = 1 Then
                    firstTable.Rows(numRows).Delete
                End If
            End If
        End If
    End If
End Sub

Sub MergeDocuments(WordApp As Object, documentsList As Variant, finalDocumentPath As String)
    Dim baseDoc As Object
    Dim sFile As String
    Dim oRng As Object
    Dim i As Integer
    
    On Error GoTo err_Handler
    
    ' Crear un nuevo documento base
    Set baseDoc = WordApp.Documents.Add
    
    ' Iterar sobre la lista de documentos a fusionar
    For i = LBound(documentsList) To UBound(documentsList)
        sFile = documentsList(i)
        
        ' Insertar el contenido del documento actual al final del documento base
        Set oRng = baseDoc.Range
        oRng.Collapse 0 ' Colapsar el rango al final del documento base
        oRng.InsertFile sFile ' Insertar el contenido del archivo actual
        
        ' Insertar un salto de página después de cada documento insertado (excepto el último)
        If i < UBound(documentsList) Then
            Set oRng = baseDoc.Range
            oRng.Collapse 0 ' Colapsar el rango al final del documento base
            'oRng.InsertBreak Type:=6 ' Insertar un salto de página
        End If
    Next i
    
    ' Guardar el archivo final
    baseDoc.SaveAs finalDocumentPath
    
    ' Cerrar el documento base
    baseDoc.Close
    
    ' Limpiar objetos
    Set baseDoc = Nothing
    Set oRng = Nothing
    
    Exit Sub
    
err_Handler:
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description
    Err.Clear
    Exit Sub
End Sub
Function EstiloExiste(docWord As Object, estilo As String) As Boolean
    Dim st As Object
    On Error Resume Next
    Set st = docWord.Styles(estilo)
    EstiloExiste = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Sub CrearEstilo(docWord As Object, estilo As String)
    Dim nuevoEstilo As Object
    On Error Resume Next
    Set nuevoEstilo = docWord.Styles.Add(Name:=estilo, Type:=1) ' Tipo 1 = Estilo de párrafo
    If Err.Number <> 0 Then
        MsgBox "No se pudo crear el estilo '" & estilo & "'. " & Err.Description, vbExclamation
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Sub ExportarTablaContenidoADocumentoWord()
    Dim appWord As Object
    Dim docWord As Object
    Dim tbl As ListObject
    Dim r As ListRow
    Dim estilo As String
    Dim seccion As String
    Dim descripcion As String
    Dim imagen As String
    Dim imagenRutaCompleta As String
    Dim parrafo As Object
    Dim rng As Object
    Dim rutaBase As String
    Dim ws As Worksheet
    Dim parrafoResultados As String
    Dim shape As Object
    
    ' Obtener la hoja activa
    Set ws = ActiveSheet
    
    ' Obtener la ruta de la carpeta donde está la hoja activa
    rutaBase = ws.Parent.Path & "\"

    ' Inicializar Word
    Set appWord = CreateObject("Word.Application")
    appWord.Visible = True
    Set docWord = appWord.Documents.Add
    
    ' Definir la tabla activa de Excel
    On Error Resume Next
    Set tbl = ws.ListObjects("Tabla_pruebas_seguridad")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "La tabla 'Tabla_pruebas_seguridad' no se encuentra en la hoja activa.", vbExclamation
        GoTo CleanUp
    End If
    
    ' Procesar cada fila de la tabla
    For Each r In tbl.ListRows
        estilo = r.Range.Cells(1, tbl.ListColumns("Estilo").Index).value
        seccion = r.Range.Cells(1, tbl.ListColumns("Sección").Index).value
        descripcion = r.Range.Cells(1, tbl.ListColumns("Descripción").Index).value
        imagen = r.Range.Cells(1, tbl.ListColumns("Imágenes").Index).value
        parrafoResultados = r.Range.Cells(1, tbl.ListColumns("Párrafo de resultados").Index).value
        
        ' Obtener la ruta completa de la imagen
        imagenRutaCompleta = rutaBase & imagen

        ' Verificar y crear el estilo si no existe
        If Not EstiloExiste(docWord, estilo) Then
            CrearEstilo docWord, estilo
        End If

        ' Agregar el encabezado con el estilo correspondiente
        If Trim(seccion) <> "" Then
        With docWord.content.Paragraphs.Add
            .Range.text = seccion
            .Range.Style = docWord.Styles(estilo)
            .Range.InsertParagraphAfter
        End With
         
        End If
        
               
        ' Agregar un párrafo con la descripción
        If Trim(descripcion) <> "" Then
            With docWord.content.Paragraphs.Add
            .Range.text = descripcion
            .Range.Style = docWord.Styles("Normal") ' Aplicar un estilo predeterminado para el párrafo de descripción
            .Format.SpaceBefore = 12 ' Espacio antes del párrafo para separación
        End With
        End If
        
        docWord.content.InsertParagraphAfter
        docWord.content.Paragraphs.Last.Range.Select

        ' Agregar la imagen si existe
        If imagen <> "" Then
            ' Verificar si la imagen existe
            If Dir(imagenRutaCompleta) <> "" Then
                ' Agregar un párrafo vacío para la imagen
                docWord.content.InsertParagraphAfter
                Set rng = docWord.content.Paragraphs.Last.Range
                
                ' Insertar la imagen en el párrafo vacío
                Set shape = docWord.InlineShapes.AddPicture(Filename:=imagenRutaCompleta, LinkToFile:=False, SaveWithDocument:=True, Range:=rng)
                
                ' Centrar la imagen
                shape.Range.ParagraphFormat.Alignment = 1 ' 1 = wdAlignParagraphCenter
                
                ' Agregar un párrafo vacío después de la imagen para el caption
                docWord.content.InsertParagraphAfter
                Set rng = docWord.content.Paragraphs.Last.Range
                
                           ' Insertar el caption debajo de la imagen
                Set captionRange = rng.Duplicate
                captionRange.Select
                appWord.Selection.MoveLeft Unit:=1, Count:=1, Extend:=0 ' wdCharacter
                appWord.CaptionLabels.Add Name:="Imagen"
                appWord.Selection.InsertCaption Label:="Imagen", TitleAutoText:="InsertarTítulo1", _
                     Title:="", Position:=1 ' wdCaptionPositionBelow, ExcludeLabel:=0
                appWord.Selection.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
                
                docWord.content.InsertAfter text:=" " & seccion
                
                ' Agregar un párrafo vacío después del caption para separación
                docWord.content.InsertParagraphAfter
                
            Else
                MsgBox "La imagen '" & imagenRutaCompleta & "' no se encuentra.", vbExclamation
            End If
        End If

        ' Agregar el párrafo de resultados si no está vacío
        If Trim(parrafoResultados) <> "" Then
            With docWord.content.Paragraphs.Add
                .Range.text = parrafoResultados
                .Range.Style = docWord.Styles("Normal") ' Aplicar un estilo predeterminado para el párrafo de resultados
                .Format.SpaceBefore = 12 ' Espacio antes del párrafo para separación
            End With
        End If
    Next r

CleanUp:
    ' Limpiar
    Set appWord = Nothing
    Set docWord = Nothing
    Set tbl = Nothing
    Set rng = Nothing
    Set shape = Nothing
End Sub

Sub LimpiarColumnaReferencias()
    Dim rng As Range
    Dim cell As Range
    Dim content As String
    Dim contentArray() As String
    Dim i As Integer
    Dim temp As String
    Dim uniqueUrls As Object
    Dim uniqueArray() As String
    Dim n As Integer
    Dim newContent As String
    Dim filteredContent As String
    
    ' Selecciona el rango deseado
    Set rng = Selection
    
    ' Recorre cada celda en el rango
    For Each cell In rng
        ' Obtiene el contenido de la celda
        content = cell.value
        
        ' Sustituye comillas dobles con saltos de línea (Char 10)
        content = Replace(content, """", Chr(10))
        
        ' Comprueba si el contenido es vacío
        If content <> "" Then
            ' Convierte el contenido en un array separado por el carácter de nueva línea
            contentArray = Split(content, Chr(10))
            
            ' Inicializa el diccionario para almacenar las URL únicas
            Set uniqueUrls = CreateObject("Scripting.Dictionary")
            
            ' Agrega las URL únicas al diccionario
            For i = LBound(contentArray) To UBound(contentArray)
                If Trim(contentArray(i)) <> "" Then
                    ' Elimina espacios en blanco, Chr(10) y Chr(13) del elemento
                    contentArray(i) = Trim(Replace(contentArray(i), Chr(13), ""))
                    contentArray(i) = Replace(contentArray(i), " ", "")
                    If InStr(1, contentArray(i), "wikipedia", vbTextCompare) = 0 Then
                        If Not uniqueUrls.Exists(contentArray(i)) Then
                            uniqueUrls.Add contentArray(i), Nothing
                        End If
                    End If
                End If
            Next i
            
            ' Convertir la colección de claves en un array
            n = uniqueUrls.Count - 1
            ReDim uniqueArray(n)
            i = 0
            For Each key In uniqueUrls.keys
                uniqueArray(i) = key
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
            newContent = Join(uniqueArray, Chr(10))
            
            ' Elimina saltos de línea iniciales y finales
            newContent = Trim(newContent)
            
            ' Filtra líneas para conservar solo las que contienen "//"
            contentArray = Split(newContent, Chr(10))
            filteredContent = ""
            
            For i = LBound(contentArray) To UBound(contentArray)
                If InStr(contentArray(i), "//") > 0 Then
                    If filteredContent <> "" Then
                        filteredContent = filteredContent & Chr(10)
                    End If
                    filteredContent = filteredContent & contentArray(i)
                End If
            Next i
            
            ' Asigna el contenido filtrado a la celda
            cell.value = filteredContent
        End If
    Next cell
End Sub



Sub LeerArchivoTXT(txtFilePath As String, dataDict As Object)
    Dim fileNumber As Integer
    Dim line As String
    Dim keyValue() As String
    Dim key As String
    Dim value As String
    
    fileNumber = FreeFile
    Open txtFilePath For Input As #fileNumber
    
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, line
        ' Divide la línea en clave y valor
        keyValue = Split(line, ":")
        If UBound(keyValue) = 1 Then
            key = Trim(keyValue(0))
            value = Trim(Mid(keyValue(1), 2, Len(keyValue(1)) - 2)) ' Extrae el valor entre comillas dobles
            ' Añadir al diccionario
            dataDict(key) = value
        End If
    Loop
    
    Close #fileNumber
End Sub






Sub WordAppAlternativeReplaceParagraph(WordApp As Object, WordDoc As Object, wordToFind As String, replaceWord As String)
    Dim findInRange As Boolean
    Dim rng As Object
    
    ' Establecer el rango al contenido del documento
    Set rng = WordDoc.content
    
    ' Configurar la búsqueda
    With rng.Find
        .text = wordToFind
        .Replacement.text = replaceWord
        .Forward = True
        .Wrap = 1 ' wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    ' Reemplazar todo el texto encontrado
    rng.Find.Execute Replace:=2 ' wdReplaceAll
End Sub



Sub ReplaceFields(WordDoc As Object, replaceDic As Object)
    Dim key As Variant
    Dim findInRange As Boolean
    Dim WordApp As Object
    Dim docContent As Object
    
    ' Obtener la aplicación de Word
    Set WordApp = WordDoc.Application
    
    ' Obtener el contenido del documento
    Set docContent = WordDoc.content
    
    ' Bucle para buscar y reemplazar todas las ocurrencias en el diccionario
    For Each key In replaceDic.keys
        ' Ir al principio del documento nuevamente
        docContent.Select
        WordApp.Selection.GoTo What:=1, Which:=1, Name:="1"
        WordApp.ActiveWindow.ActivePane.View.SeekView = 0
        
        ' Configurar la búsqueda
        With WordApp.Selection.Find
            .text = key
            .Forward = True
            .Wrap = 1 ' wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            
            ' Intentar encontrar y reemplazar en el cuerpo del documento
            Do
                findInRange = .Execute
                If findInRange Then
                    ' Realizar el reemplazo
                    WordApp.Selection.text = CStr(replaceDic(key))
                    ' Ir al principio del documento nuevamente
                    WordApp.Selection.GoTo What:=1, Which:=1, Name:="1"
                End If
            Loop While findInRange
        End With
    Next key

    ' Limpiar
    Set docContent = Nothing
    Set WordApp = Nothing
End Sub


Sub GenerarReportesVulnsAppsINAI()
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim plantillaReportePath As String
    Dim plantillaReportePath2 As String
    Dim plantillaVulnerabilidadesPath As String
    Dim carpetaSalida As String
    Dim archivoTemp As String
    Dim fileSystem As Object
    Dim dlg As FileDialog
    Dim rng As Range
    Dim replaceDic As Object
    Dim cell As Range
    Dim rowCount As Integer
    Dim i As Integer
    Dim tempFolder As String
    Dim tempFolderGenerados As String
    Dim finalDocumentPath As String
    Dim tempDocPath As String
    Dim tempDocVulnerabilidadesPath As String
    Dim tempFileName As String
    Dim secVulnerabilidades As String
    Dim rngReplace As Object
    Dim campoArchivoPath As String
    Dim campoLine As String
    Dim partes() As String
    Dim key As String
    Dim value As String
    Dim documentsList() As String
    Dim numDocuments As Integer
    Dim severityCounts As Object
    Dim severity As Variant
    Dim totalVulnerabilidades As Integer
    Dim chart As Object
    Dim seriesCollection As Object
    Dim dataLabels As Object
    Dim severidadColumna As Integer
    Dim countBAJA As Integer
    Dim countMEDIA As Integer
    Dim countALTA As Integer
    Dim countCRITICAS As Integer
    Dim ils As Object
    Dim wb As Object
    Dim SourceSheet As Object
    Dim appName As String ' Variable para el nombre de la aplicación

    ' Crear un diálogo para seleccionar el archivo CSV
    campoArchivoPath = Application.GetOpenFilename("Archivos CSV (*.csv), *.csv", , "Seleccionar archivo CSV")
    If campoArchivoPath = "Falso" Then
        MsgBox "No se seleccionó ningún archivo CSV. La macro se detendrá."
        Exit Sub
    End If

    ' Leer campos de reemplazo desde el archivo CSV
    Set replaceDic = CreateObject("Scripting.Dictionary")
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set ts = fileSystem.OpenTextFile(campoArchivoPath, 1, False, 0)
    
    ' Leer el archivo línea por línea
    Do Until ts.AtEndOfStream
        csvLine = ts.ReadLine
        
        ' Separar la línea en clave y valor usando la coma como delimitador
        partes = Split(csvLine, ",", 2) ' Divide en dos partes (clave, valor)
        
        If UBound(partes) = 1 Then
            key = Trim(partes(0))
            value = Trim(partes(1))
            
            ' Añadir al diccionario
            replaceDic(key) = value
        End If
    Loop
    ts.Close
    
    ' Extraer el nombre de la aplicación del diccionario
    If replaceDic.Exists("«Aplicación»") Then
        appName = replaceDic("«Aplicación»")
    Else
        MsgBox "No se encontró el campo 'Aplicación' en el archivo CSV.", vbExclamation
        Exit Sub
    End If

    ' Crear diálogos para seleccionar plantillas y carpeta de salida
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    dlg.Title = "Seleccionar la plantilla de reporte técnico"
    dlg.Filters.Clear
    dlg.Filters.Add "Archivos de Word", "*.docx"
    
    If dlg.Show = -1 Then
        plantillaReportePath = dlg.SelectedItems(1)
    Else
        MsgBox "No se seleccionó ningún archivo. La macro se detendrá."
        Exit Sub
    End If
    
    dlg.Title = "Seleccionar la plantilla de reporte ejecutivo"
    
    If dlg.Show = -1 Then
        plantillaReportePath2 = dlg.SelectedItems(1)
    Else
        MsgBox "No se seleccionó ningún archivo. La macro se detendrá."
        Exit Sub
    End If
    
    dlg.Title = "Seleccionar la plantilla de tabla de vulnerabilidades"
    
    If dlg.Show = -1 Then
        plantillaVulnerabilidadesPath = dlg.SelectedItems(1)
    Else
        MsgBox "No se seleccionó ningún archivo. La macro se detendrá."
        Exit Sub
    End If
    
    ' Solicitar la carpeta de salida
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccionar Carpeta de Salida"
        If .Show = -1 Then
            carpetaSalida = .SelectedItems(1)
        Else
            MsgBox "No se seleccionó ninguna carpeta. La macro se detendrá."
            Exit Sub
        End If
    End With
    
    ' Crear una subcarpeta con el nombre de la aplicación
    carpetaSalida = carpetaSalida & "\AV " & appName
    On Error Resume Next
    MkDir carpetaSalida
    On Error GoTo 0
    
    ' Crear una instancia de Word
    On Error Resume Next
    Set WordApp = CreateObject("Word.Application")
    On Error GoTo 0
    
    If WordApp Is Nothing Then
        MsgBox "No se puede iniciar Microsoft Word."
        Exit Sub
    End If
    
    ' Crear una carpeta temporal
    tempFolder = Environ("TEMP") & "\tmp-" & Format(Now(), "yyyymmddhhmmss")
    On Error Resume Next
    MkDir tempFolder
    On Error GoTo 0
    
    ' Crear una carpeta para documentos generados
    tempFolderGenerados = tempFolder & "\Documentos generados"
    On Error Resume Next
    MkDir tempFolderGenerados
    On Error GoTo 0
    
    ' Crear y abrir documentos de reporte técnico y ejecutivo
    For Each plantilla In Array(plantillaReportePath, plantillaReportePath2)
        archivoTemp = tempFolder & "\" & fileSystem.GetFileName(plantilla)
        fileSystem.CopyFile plantilla, archivoTemp
        
        Set WordDoc = WordApp.Documents.Open(archivoTemp)
        WordApp.Visible = False
        
        ReplaceFields WordDoc, replaceDic
        
        If plantilla = plantillaReportePath Then
            tempDocPath = tempFolder & "\SSIFO14-03 Informe Técnico.docx"
            WordDoc.SaveAs2 tempDocPath
        Else
            tempDocPath2 = tempFolder & "\SSIFO15-03 Informe Ejecutivo.docx"
            WordDoc.SaveAs2 tempDocPath2
        End If
        
        WordDoc.Close False
    Next
    
    ' Solicita al usuario seleccionar el rango de celdas
    On Error Resume Next
    Set selectedRange = Application.InputBox("Seleccione el rango de celdas que contienen las columnas a considerar", Type:=8)
    On Error GoTo 0
    
    If selectedRange Is Nothing Then
        MsgBox "No se ha seleccionado un rango válido.", vbExclamation
        Exit Sub
    End If
    
    ' Verifica si el rango seleccionado está dentro de una tabla
    On Error Resume Next
    Set rng = selectedRange.ListObject.Range
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "El rango seleccionado no está dentro de una tabla.", vbExclamation
        Exit Sub
    End If

    ' Inicializar el diccionario para contar severidades
    Set severityCounts = CreateObject("Scripting.Dictionary")
    
    ' Buscar la columna "Severidad"
    severidadColumna = -1
    For i = 1 To rng.Columns.Count
        If rng.Cells(1, i).value = "Severidad" Then
            severidadColumna = i
            Exit For
        End If
    Next i
    
    If severidadColumna = -1 Then
        MsgBox "No se encontró la columna 'Severidad' en el rango seleccionado.", vbExclamation
        Exit Sub
    End If
    
    ' Contar severidades
    For i = 2 To rng.Rows.Count
        severity = rng.Cells(i, severidadColumna).value
        If severity <> "" Then
            If severityCounts.Exists(severity) Then
                severityCounts(severity) = severityCounts(severity) + 1
            Else
                severityCounts.Add severity, 1
            End If
        End If
    Next i
    
    ' Inicializar el diccionario para contar tipos vulnerabilidades
    Set vulntypesCounts = CreateObject("Scripting.Dictionary")
    
    ' Buscar la columna "Tipo de vulnerabilidad"
    tiposvulnerabilidadColumna = -1
    For i = 1 To rng.Columns.Count
        If rng.Cells(1, i).value = "Tipo de vulnerabilidad" Then
            tiposvulnerabilidadColumna = i
            Exit For
        End If
    Next i
    
    If tiposvulnerabilidadColumna = -1 Then
        MsgBox "No se encontró la columna 'Tipo de vulnerabilidad' en el rango seleccionado.", vbExclamation
        Exit Sub
    End If
    
    ' Contar tipos vulnerabilidades
    For i = 2 To rng.Rows.Count
        vulntypes = rng.Cells(i, tiposvulnerabilidadColumna).value
        If vulntypes <> "" Then
            If vulntypesCounts.Exists(vulntypes) Then
                vulntypesCounts(vulntypes) = vulntypesCounts(vulntypes) + 1
            Else
                vulntypesCounts.Add vulntypes, 1
            End If
        End If
    Next i

    ' Inicializar conteos
    countBAJA = IIf(severityCounts.Exists("BAJA"), severityCounts("BAJA"), 0)
    countMEDIA = IIf(severityCounts.Exists("MEDIA"), severityCounts("MEDIA"), 0)
    countALTA = IIf(severityCounts.Exists("ALTA"), severityCounts("ALTA"), 0)
    countCRITICAS = IIf(severityCounts.Exists("CRÍTICAS"), severityCounts("CRÍTICAS"), 0)

    ' Calcular total de vulnerabilidades
    totalVulnerabilidades = countBAJA + countMEDIA + countALTA + countCRITICAS

    ' Copia la plantilla de vulnerabilidades
    tempDocVulnerabilidadesPath = tempFolder & "\Plantilla_Vulnerabilidades.docx"
    fileSystem.CopyFile plantillaVulnerabilidadesPath, tempDocVulnerabilidadesPath
    
    ' Genera documentos de Word por cada registro
    rowCount = rng.Rows.Count
    For i = 2 To rowCount
        Set replaceDic = CreateObject("Scripting.Dictionary")
        
        For Each cell In selectedRange.Rows(1).Cells
            replaceDic("«" & cell.value & "»") = rng.Cells(i, cell.Column).value
        Next cell
        
        tempFileName = "Documento_" & i & ".docx"
        fileSystem.CopyFile tempDocVulnerabilidadesPath, tempFolderGenerados & "\" & tempFileName
        
        On Error Resume Next
        Set WordDoc = WordApp.Documents.Open(tempFolderGenerados & "\" & tempFileName)
        If Err.Number <> 0 Then
            MsgBox "No se pudo abrir el archivo: " & tempFolderGenerados & "\" & tempFileName, vbCritical
            Err.Clear
            Exit Sub
        End If
        On Error GoTo 0
        
        ReplaceFields WordDoc, replaceDic
        FormatRiskLevelCell WordDoc.Tables(1).cell(1, 2)
        EliminarUltimasFilasSiEsSalidaPruebaSeguridad WordDoc, replaceDic
        WordDoc.Save
        WordDoc.Close
        
        numDocuments = numDocuments + 1
        ReDim Preserve documentsList(numDocuments - 1)
        documentsList(numDocuments - 1) = tempFolderGenerados & "\" & tempFileName
    Next i
    
    ' Combina todos los archivos en uno solo
    finalDocumentPath = tempFolder & "\Tablas_vulnerabilidades.docx"
    MergeDocuments WordApp, documentsList, finalDocumentPath
    
    ' Actualizar el documento de reporte técnico
    Set WordDoc = WordApp.Documents.Open(tempDocPath)
    secVulnerabilidades = "{{Sección de tablas de vulnerabilidades}}"
    
    Set rngReplace = WordDoc.content
    rngReplace.Find.ClearFormatting
    With rngReplace.Find
        .text = secVulnerabilidades
        .Replacement.text = ""
        .Forward = True
        .Wrap = 1 ' wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    If rngReplace.Find.Execute Then
        rngReplace.text = ""
        rngReplace.InsertFile finalDocumentPath
    End If
    
    ' Reemplazar el total de vulnerabilidades
    Set rngReplace = WordDoc.content
    rngReplace.Find.ClearFormatting
    With rngReplace.Find
        .text = "«Total de vulnerabilidades»"
        .Replacement.text = totalVulnerabilidades
        .Forward = True
        .Wrap = 1 ' wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    If rngReplace.Find.Execute Then
        rngReplace.text = totalVulnerabilidades
    End If

    ' Actualizar el gráfico InlineShape número 1
    On Error Resume Next
    Set ils = WordDoc.InlineShapes(1)
    If Err.Number = 0 Then
        On Error GoTo 0
        
        If ils.Type = 12 Then ' Verificar si es un gráfico
            Set chart = ils.chart
            chart.ChartData.Activate
            Set wb = chart.ChartData.Workbook
            Set SourceSheet = wb.Sheets(1)
            
            ' Actualizar los datos del gráfico
            SourceSheet.Cells(2, 2).Value2 = countBAJA
            SourceSheet.Cells(3, 2).Value2 = countMEDIA
            SourceSheet.Cells(4, 2).Value2 = countALTA
            SourceSheet.Cells(5, 2).Value2 = countCRITICAS
            
            wb.Close
        Else
            MsgBox "El InlineShape número 1 no es un gráfico."
        End If
    Else
        MsgBox "No se encontró el gráfico InlineShape número 1."
    End If
    On Error GoTo 0

      ' Actualizar todos los gráficos en el documento
    For i = 1 To WordDoc.InlineShapes.Count
        Set ils = WordDoc.InlineShapes(i)
        If ils.Type = wdInlineShapeChart Then ' Verificar si el InlineShape es un gráfico
            Set chart = ils.chart
            If Not chart Is Nothing Then
                chart.ChartData.Activate
                ' Intentar obtener el libro de trabajo asociado y cerrarlo
                Set wb = chart.ChartData.Workbook
                If Not wb Is Nothing Then
                    wb.Application.Visible = False
                    On Error Resume Next
                    wb.Close SaveChanges:=False
                    On Error GoTo 0
                End If
                ' Refrescar el gráfico
                chart.Refresh
            End If
        End If
    Next i
    On Error GoTo 0

    ' Guardar el documento de reporte técnico final en la subcarpeta
    WordDoc.SaveAs2 carpetaSalida & "\SSIFO14-03 Informe Técnico.docx"
    WordDoc.Close False

    ' Actualizar el documento de reporte ejecutivo
    Set WordDoc = WordApp.Documents.Open(tempDocPath2)
    
    Set rngReplace = WordDoc.content
    rngReplace.Find.ClearFormatting
    With rngReplace.Find
        .text = "«Total de vulnerabilidades»"
        .Replacement.text = totalVulnerabilidades
        .Forward = True
        .Wrap = 1 ' wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    If rngReplace.Find.Execute Then
        rngReplace.text = totalVulnerabilidades
    End If

    ' Actualizar el gráfico InlineShape número 1
    On Error Resume Next
    Set ils = WordDoc.InlineShapes(1)
    If Err.Number = 0 Then
        On Error GoTo 0
        
        If ils.Type = 12 Then ' Verificar si es un gráfico
            Set chart = ils.chart
            chart.ChartData.Activate
            Set wb = chart.ChartData.Workbook
            Set SourceSheet = wb.Sheets(1)
            
            ' Actualizar los datos del gráfico
            SourceSheet.Cells(2, 2).Value2 = countBAJA
            SourceSheet.Cells(3, 2).Value2 = countMEDIA
            SourceSheet.Cells(4, 2).Value2 = countALTA
            SourceSheet.Cells(5, 2).Value2 = countCRITICAS
            
            wb.Close SaveChanges:=True
            chart.Refresh
        Else
            MsgBox "El InlineShape número 1 no es un gráfico."
        End If
    Else
        MsgBox "No se encontró el gráfico InlineShape número 1."
    End If
    On Error GoTo 0

    ' Actualizar el gráfico InlineShape número 2
    On Error Resume Next
    Set ils = WordDoc.InlineShapes(2)
    If Err.Number = 0 Then
        On Error GoTo 0
        
        If ils.Type = 12 Then ' Verificar si es un gráfico
            Set chart = ils.chart
            chart.ChartData.Activate
            Set wb = chart.ChartData.Workbook
            Set SourceSheet = wb.Sheets(1)
            
            ' Actualizar los datos del gráfico
            SourceSheet.Cells(2, 2).Value2 = countBAJA
            SourceSheet.Cells(3, 2).Value2 = countMEDIA
            SourceSheet.Cells(4, 2).Value2 = countALTA
            SourceSheet.Cells(5, 2).Value2 = countCRITICAS
            
            wb.Close
            chart.Refresh
        Else
            MsgBox "El InlineShape número 2 no es un gráfico."
        End If
    Else
        MsgBox "No se encontró el gráfico InlineShape número 2."
    End If
    On Error GoTo 0

      ' Actualizar todos los gráficos en el documento
    For i = 1 To WordDoc.InlineShapes.Count
        Set ils = WordDoc.InlineShapes(i)
        If ils.Type = wdInlineShapeChart Then ' Verificar si el InlineShape es un gráfico
            Set chart = ils.chart
            If Not chart Is Nothing Then
                chart.ChartData.Activate
                ' Intentar obtener el libro de trabajo asociado y cerrarlo
                Set wb = chart.ChartData.Workbook
                If Not wb Is Nothing Then
                    wb.Application.Visible = False
                    On Error Resume Next
                    wb.Close SaveChanges:=False
                    On Error GoTo 0
                End If
                ' Refrescar el gráfico
                chart.Refresh
            End If
        End If
    Next i
    On Error GoTo 0
    

    ' Guardar el documento de reporte ejecutivo final en la subcarpeta
    WordDoc.SaveAs2 carpetaSalida & "\SSIFO15-03 Informe Ejecutivo.docx"
    WordDoc.Close False
    
    ' Mover los documentos generados y el archivo consolidado a la subcarpeta de salida
    fileSystem.MoveFolder tempFolderGenerados, carpetaSalida & "\Documentos generados"
    fileSystem.MoveFile finalDocumentPath, carpetaSalida & "\Tablas_vulnerabilidades.docx"
    
    ' Cerrar la aplicación de Word
    WordApp.Quit
    Set WordDoc = Nothing
    Set WordApp = Nothing
    Set fileSystem = Nothing
    
    ' Mostrar mensaje de éxito
    MsgBox "Se han generado los documentos de Word correctamente.", vbInformation
End Sub
