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
            For Each key In uniqueUrls.Keys
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

' Posible divisor de comentarios
Sub LimpiarSalida()
    Dim rng As Range
    Dim c As Range
    Dim celda As Range
    Dim lineas As Variant
    Dim i As Integer
    
    ' Asume que el rango seleccionado es el que quieres modificar
    Set rng = Selection
    
    ' Primero, quita los espacios en blanco y caracteres de tabulación de las celdas
    For Each c In rng
        ' Reemplazar caracteres de tabulación con espacios
        c.value = Replace(c.value, Chr(9), " ")
        ' Quitar espacios en blanco adicionales
        c.value = Application.Trim(c.value)
    Next c

    ' Luego, elimina las líneas vacías y los saltos de línea finales de las celdas
    For Each celda In rng
        ' Verificar si la celda tiene texto
        If Not IsEmpty(celda.value) Then
            ' Reemplazar diferentes saltos de línea con vbLf
            Dim contenido As String
            contenido = Replace(Replace(Replace(celda.value, vbCrLf, vbLf), vbCr, vbLf), vbLf & vbLf, vbLf)
            
            ' Si el contenido comienza con vbLf, quitarlo
            If Left(contenido, 1) = vbLf Then
                contenido = Mid(contenido, 2)
            End If
            
            ' Si el contenido termina con vbLf, quitarlo
            If Right(contenido, 1) = vbLf Then
                contenido = Left(contenido, Len(contenido) - 1)
            End If
            
            ' Dividir el contenido de la celda en un array de líneas
            lineas = Split(contenido, vbLf)
            
            ' Iterar sobre cada línea del array
            For i = LBound(lineas) To UBound(lineas)
                ' Verificar si la línea está vacía y eliminarla
                If Trim(lineas(i)) = "" Then
                    lineas(i) = vbNullString
                End If
            Next i
            
            ' Unir el array de líneas de nuevo en una cadena y asignarlo a la celda
            ' Además, eliminar posibles saltos de línea al final del contenido
            celda.value = Join(lineas, vbLf)
            ' Eliminar saltos de línea al final
            If Right(celda.value, 1) = vbLf Then
                celda.value = Left(celda.value, Len(celda.value) - 1)
            End If
        End If
    Next celda
End Sub

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
        .pattern = "([^.()\r\n-])(?![^(]*\)|[-])[^\S\r\n]*[\r\n]+" ' Expresión regular para encontrar saltos de línea o saltos de carro sin un punto antes y no seguidos de paréntesis ni de guión
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
            For Each key In uniqueUrls.Keys
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
    For Each key In replaceDic.Keys
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
                    WordApp.Selection.text = replaceDic(key)
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
    Dim campoData() As String
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
    
    
    ' Crear un diálogo para seleccionar el archivo de campos de reemplazo
    campoArchivoPath = Application.GetOpenFilename("Archivos de Texto (*.txt), *.txt", , "Seleccionar archivo de campos de reemplazo")
    If campoArchivoPath = "Falso" Then
        MsgBox "No se seleccionó ningún archivo de campos de reemplazo. La macro se detendrá."
        Exit Sub
    End If

    ' Leer campos de reemplazo desde el archivo de texto
    Set replaceDic = CreateObject("Scripting.Dictionary")
    Open campoArchivoPath For Input As #1
    Do While Not EOF(1)
        Line Input #1, campoLine
        
        ' Encontrar la primera aparición de ":"
        colonPos = InStr(campoLine, ":")
        
        If colonPos > 0 Then
            ' Extraer la clave y el valor basado en el primer ":"
            key = Trim(Left(campoLine, colonPos - 1))
            value = Trim(Mid(campoLine, colonPos + 1))
            
            ' Extraer el texto entre comillas dobles, si existe
            If Len(value) > 0 And Left(value, 1) = Chr(34) Then
                ' Buscar la posición de la primera y última comilla doble
                startPos = InStr(value, Chr(34))
                endPos = InStrRev(value, Chr(34))
                ' Extraer el texto entre las comillas dobles
                If startPos > 0 And endPos > startPos Then
                    value = Mid(value, startPos + 1, endPos - startPos - 1)
                End If
            End If
            
            replaceDic(key) = value
        End If
    Loop
    Close #1

    ' Crear un diálogo para seleccionar la plantilla de reporte técnico
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
    
    
    ' Crear un diálogo para seleccionar la plantilla de reporte ejecutivo
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    dlg.Title = "Seleccionar la plantilla de reporte técnico"
    dlg.Filters.Clear
    dlg.Filters.Add "Archivos de Word", "*.docx"
    
    If dlg.Show = -1 Then
        plantillaReportePath2 = dlg.SelectedItems(1)
    Else
        MsgBox "No se seleccionó ningún archivo. La macro se detendrá."
        Exit Sub
    End If
    
    ' Crear un diálogo para seleccionar la plantilla de tabla de vulnerabilidades
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
    
    ' Crear una instancia de Word
    On Error Resume Next
    Set WordApp = CreateObject("Word.Application")
    On Error GoTo 0
    
    If WordApp Is Nothing Then
        MsgBox "No se puede iniciar Microsoft Word."
        Exit Sub
    End If
    
    ' Crear una carpeta temporal en la carpeta de archivos temporales del sistema
    tempFolder = Environ("TEMP") & "\tmp-" & Format(Now(), "yyyymmddhhmmss")
    On Error Resume Next
    MkDir tempFolder
    On Error GoTo 0
    
    ' Crear una carpeta para los documentos generados
    tempFolderGenerados = tempFolder & "\DocumentosGenerados"
    On Error Resume Next
    MkDir tempFolderGenerados
    On Error GoTo 0
    
    ' Crear una copia temporal del archivo de plantilla de reporte técnico
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    archivoTemp = tempFolder & "\" & fileSystem.GetFileName(plantillaReportePath)
    fileSystem.CopyFile plantillaReportePath, archivoTemp
    
    ' Abrir la plantilla de reporte técnico en Word
    Set WordDoc = WordApp.Documents.Open(archivoTemp)
    WordApp.Visible = False
    
    ' Reemplazar los campos en la plantilla de reporte técnico
    ReplaceFields WordDoc, replaceDic
    
    ' Guardar el documento de reporte técnico temporalmente
    tempDocPath = tempFolder & "\SSIFO14-03 Informe Técnico.docx"
    WordDoc.SaveAs2 tempDocPath
    WordDoc.Close False

      ' Crear una copia temporal del archivo de plantilla de reporte ejecutivo
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    archivoTemp = tempFolder & "\" & fileSystem.GetFileName(plantillaReportePath2)
    fileSystem.CopyFile plantillaReportePath2, archivoTemp
    
    ' Abrir la plantilla de reporte ejecutivo en Word
    Set WordDoc = WordApp.Documents.Open(archivoTemp)
    WordApp.Visible = False
    
    ' Reemplazar los campos en la plantilla de reporte ejecutivo
    ReplaceFields WordDoc, replaceDic
    
    ' Guardar el documento de reporte ejecutivo temporalmente
    tempDocPath = tempFolder & "\SSIFO15-03 Informe ejecutivo.docx"
    WordDoc.SaveAs2 tempDocPath2
    WordDoc.Close False
    
    ' Solicita al usuario seleccionar el rango de celdas que contienen las columnas a considerar
    On Error Resume Next
    Set selectedRange = Application.InputBox("Seleccione el rango de celdas que contienen las columnas a considerar", Type:=8)
    On Error GoTo 0
    
    
    ' Solicita al usuario seleccionar el rango de celdas que contienen las columnas a considerar
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
    
    ' Buscar la columna "Severidad" en la primera fila para identificar su índice
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
    
    ' Recorre la columna "Severidad" y cuenta las ocurrencias
    For i = 2 To rng.Rows.Count ' Empieza desde la segunda fila, asumiendo que la primera fila es el encabezado
        severity = rng.Cells(i, severidadColumna).value
        If severity <> "" Then
            If severityCounts.Exists(severity) Then
                severityCounts(severity) = severityCounts(severity) + 1
            Else
                severityCounts.Add severity, 1
            End If
        End If
    Next i
    
    ' Inicializar los conteos de severidad
    countBAJA = IIf(severityCounts.Exists("BAJA"), severityCounts("BAJA"), 0)
    countMEDIA = IIf(severityCounts.Exists("MEDIA"), severityCounts("MEDIA"), 0)
    countALTA = IIf(severityCounts.Exists("ALTA"), severityCounts("ALTA"), 0)
    countCRITICAS = IIf(severityCounts.Exists("CRÍTICAS"), severityCounts("CRÍTICAS"), 0)
    
    ' Calcular el total de vulnerabilidades
    totalVulnerabilidades = countBAJA + countMEDIA + countALTA + countCRITICAS

    ' Copia la plantilla de vulnerabilidades a la carpeta temporal
    tempDocVulnerabilidadesPath = tempFolder & "\Plantilla_Vulnerabilidades.docx"
    fileSystem.CopyFile plantillaVulnerabilidadesPath, tempDocVulnerabilidadesPath
    
    ' Inicializar el arreglo de documentos
    numDocuments = 0
    ReDim documentsList(numDocuments)
    
    ' Genera un archivo de Word por cada registro de la tabla
    rowCount = rng.Rows.Count
    For i = 2 To rowCount ' Empezamos desde la segunda fila para los datos reales
        ' Crear un nuevo diccionario para cada fila
        Set replaceDic = CreateObject("Scripting.Dictionary")
        
        ' Llena el diccionario de reemplazo con los datos de la fila actual de la tabla de Excel
        For Each cell In selectedRange.Rows(1).Cells ' Tomamos la primera fila para los nombres de los campos
            replaceDic("«" & cell.value & "»") = rng.Cells(i, cell.Column).value
        Next cell
        
        ' Crea una copia del documento de Word en la carpeta temporal
        tempFileName = "Documento_" & i & ".docx"
        fileSystem.CopyFile tempDocVulnerabilidadesPath, tempFolderGenerados & "\" & tempFileName
        ' Abre la copia del documento de Word
        On Error Resume Next
        Set WordDoc = WordApp.Documents.Open(tempFolderGenerados & "\" & tempFileName)
        If Err.Number <> 0 Then
            MsgBox "No se pudo abrir el archivo: " & tempFolderGenerados & "\" & tempFileName, vbCritical
            Err.Clear
            Exit Sub
        End If
        On Error GoTo 0
        
        ' Realiza los reemplazos en el documento de Word
        ReplaceFields WordDoc, replaceDic
        FormatRiskLevelCell WordDoc.Tables(1).cell(1, 2)
        ' Guarda y cierra el documento de Word
        ' Antes de guardar el documento de Word
        EliminarUltimasFilasSiEsSalidaPruebaSeguridad WordDoc, replaceDic
        WordDoc.Save
        WordDoc.Close
        
        ' Agregar el documento generado a la lista
        numDocuments = numDocuments + 1
        ReDim Preserve documentsList(numDocuments - 1)
        documentsList(numDocuments - 1) = tempFolderGenerados & "\" & tempFileName
    Next i
    
    ' Combina todos los archivos en uno solo
    finalDocumentPath = tempFolder & "\Documento_Consolidado.docx"
    MergeDocuments WordApp, documentsList, finalDocumentPath
    
    ' Abrir el documento de reporte técnico
    Set WordDoc = WordApp.Documents.Open(tempDocPath)
    
    ' Reemplazar el texto específico por el contenido del archivo consolidado
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
    
    ' Ejecutar búsqueda y reemplazo
    If rngReplace.Find.Execute Then
        ' Insertar el contenido del documento consolidado
        rngReplace.text = ""
        rngReplace.InsertFile finalDocumentPath
    End If
    
    ' Reemplazar el total de vulnerabilidades en el documento
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
    
    ' Ejecutar búsqueda y reemplazo
    If rngReplace.Find.Execute Then
        rngReplace.text = totalVulnerabilidades
    End If
    
    ' Actualizar el gráfico InlineShape número 1
    On Error Resume Next
    Set chart = WordDoc.InlineShapes(1).OLEFormat.Object.chart
    If Err.Number = 0 Then
        ' Activar el gráfico para poder editar los datos
        chart.ChartData.Activate
        Dim wb As Object
        Dim SourceSheet As Object
        Set wb = chart.ChartData.Workbook
        Set SourceSheet = wb.ActiveSheet
        
        ' Actualizar los datos del gráfico con los valores calculados
        SourceSheet.Range("B2").Value2 = countBAJA ' Actualizar el número de vulnerabilidades BAJA
        SourceSheet.Range("B3").Value2 = countMEDIA ' Actualizar el número de vulnerabilidades MEDIA
        SourceSheet.Range("B4").Value2 = countALTA ' Actualizar el número de vulnerabilidades ALTA
        SourceSheet.Range("B5").Value2 = countCRITICAS ' Actualizar el número de vulnerabilidades CRÍTICAS
        
        ' Cerrar el libro de datos
        wb.Close SaveChanges:=True
        
        ' Actualizar el gráfico
        chart.Refresh
        
        Set seriesCollection = chart.seriesCollection(1)
        seriesCollection.HasDataLabels = True
        Set dataLabels = seriesCollection.dataLabels
        dataLabels.ShowValue = True
    Else
        MsgBox "No se encontró el gráfico InlineShape número 1."
    End If
    On Error GoTo 0
    
    ' Guardar el documento de reporte técnico final
    WordDoc.SaveAs2 carpetaSalida & "\SSIFO14-03 Informe Técnico.docx"
    WordDoc.Close False
    
    
    ' Abrir el documento de reporte técnico
    Set WordDoc = WordApp.Documents.Open(tempDocPath2)
    

    ' Reemplazar el total de vulnerabilidades en el documento
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
    
    ' Ejecutar búsqueda y reemplazo
    If rngReplace.Find.Execute Then
        rngReplace.text = totalVulnerabilidades
    End If
    
    ' Actualizar el gráfico InlineShape número 1
    On Error Resume Next
    Set chart = WordDoc.InlineShapes(1).OLEFormat.Object.chart
    If Err.Number = 0 Then
        ' Activar el gráfico para poder editar los datos
        chart.ChartData.Activate
        Dim wb As Object
        Dim SourceSheet As Object
        Set wb = chart.ChartData.Workbook
        Set SourceSheet = wb.ActiveSheet
        
        ' Actualizar los datos del gráfico con los valores calculados
        SourceSheet.Range("B2").Value2 = countBAJA ' Actualizar el número de vulnerabilidades BAJA
        SourceSheet.Range("B3").Value2 = countMEDIA ' Actualizar el número de vulnerabilidades MEDIA
        SourceSheet.Range("B4").Value2 = countALTA ' Actualizar el número de vulnerabilidades ALTA
        SourceSheet.Range("B5").Value2 = countCRITICAS ' Actualizar el número de vulnerabilidades CRÍTICAS
        
        ' Cerrar el libro de datos
        wb.Close SaveChanges:=True
        
        ' Actualizar el gráfico
        chart.Refresh
        
        Set seriesCollection = chart.seriesCollection(1)
        seriesCollection.HasDataLabels = True
        Set dataLabels = seriesCollection.dataLabels
        dataLabels.ShowValue = True
    Else
        MsgBox "No se encontró el gráfico InlineShape número 1."
    End If
    On Error GoTo 0
    
    ' Guardar el documento de reporte técnico final
    WordDoc.SaveAs2 carpetaSalida & "\SSIFO15-03 Informe Ejecutivo.docx"
    WordDoc.Close False
    
    ' Mueve los documentos generados a la carpeta de salida final
    fileSystem.MoveFolder tempFolderGenerados, carpetaSalida & "\DocumentosGenerados"
    
    ' Mueve el archivo consolidado a la carpeta de salida final
    fileSystem.MoveFile finalDocumentPath, carpetaSalida & "\Documento_Consolidado.docx"
    
    ' Cerrar la aplicación de Word
    WordApp.Quit
    Set WordDoc = Nothing
    Set WordApp = Nothing
    Set fileSystem = Nothing
    
    ' Muestra un mensaje de éxito
    MsgBox "Se han generado los documentos de Word correctamente.", vbInformation
End Sub



