Attribute VB_Name = "ExcelModulosCibersecurity"
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

Sub ReplaceWithURLs()
    Dim cell As Range
    Dim parts As Variant
    Dim url As String
    Dim i As Integer
    
    ' Recorre cada celda en el rango seleccionado
    For Each cell In Selection
        If cell.Value <> "" Then
            ' Separa la cadena por comas
            parts = Split(cell.Value, ",")
            
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
            cell.Value = url
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
    Dim texto As String
    Dim primeraLetra As String
    Dim restoTexto As String

    ' Recorre cada celda en el rango seleccionado
    For Each celda In Selection
        If Not IsEmpty(celda.Value) Then
            texto = celda.Value
            ' Convierte todo el texto a minúsculas
            texto = LCase(texto)
            ' Extrae la primera letra
            primeraLetra = UCase(Left(texto, 1))
            ' Extrae el resto del texto
            restoTexto = Mid(texto, 2)
            ' Combina la primera letra en mayúsculas con el resto del texto en minúsculas
            celda.Value = primeraLetra & restoTexto
        End If
    Next celda
End Sub


Sub QuitarEspacios()
    Dim rng As Range
    Dim c As Range
    
    Set rng = Selection 'asume que el rango seleccionado es el que quieres modificar
    
    For Each c In rng 'recorre cada celda del rango
        c.Value = Application.Trim(c.Value) 'quita los espacios de la celda
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
        c.Value = Replace(c.Value, Chr(9), " ")
        ' Quitar espacios en blanco adicionales
        c.Value = Application.Trim(c.Value)
    Next c

    ' Luego, elimina las líneas vacías y los saltos de línea finales de las celdas
    For Each celda In rng
        ' Verificar si la celda tiene texto
        If Not IsEmpty(celda.Value) Then
            ' Reemplazar diferentes saltos de línea con vbLf
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
            celda.Value = Join(lineas, vbLf)
            ' Eliminar saltos de línea al final
            If Right(celda.Value, 1) = vbLf Then
                celda.Value = Left(celda.Value, Len(celda.Value) - 1)
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
                .SortFields.Add Key:=ws.ListObjects(tabla.Name).ListColumns("Severidad").Range, _
                    SortOn:=xlSortOnCellColor, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal
                .SortFields(1).SortOnValue.Color = RGB(0, 176, 80)
                .Apply
                
                ' Limpiar campos de ordenación previos
                .SortFields.Clear
                
                ' Ordenar por color de relleno amarillo
                .SortFields.Add Key:=ws.ListObjects(tabla.Name).ListColumns("Severidad").Range, _
                    SortOn:=xlSortOnCellColor, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal
                .SortFields(1).SortOnValue.Color = RGB(255, 255, 0)
                .Apply
                
                ' Limpiar campos de ordenación previos
                .SortFields.Clear
                
                ' Ordenar por color de relleno rojo
                .SortFields.Add Key:=ws.ListObjects(tabla.Name).ListColumns("Severidad").Range, _
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

Sub GenerarDocumentosWordVulnerabilidades()
    Dim rng As Range
    Dim tbl As ListObject
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim templatePath As String
    Dim outputPath As String
    Dim replaceDic As Object
    Dim cell As Range
    Dim colIndex As Integer
    Dim rowCount As Integer
    Dim i As Integer
    Dim tempFolder As String
    Dim tempFolderPath As String
    Dim saveFolder As String
    Dim selectedRange As Range ' Variable para almacenar el rango seleccionado por el usuario
    Dim documentsList() As String ' Lista para almacenar los documentos generados
    
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
    
    ' Solicita al usuario la ruta del documento de Word
    templatePath = Application.GetOpenFilename("Documentos de Word (*.docx), *.docx", , "Seleccione un documento de Word como plantilla")
    If templatePath = "Falso" Then Exit Sub
    
    ' Solicita al usuario la carpeta donde desea guardar los archivos generados
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleccione la carpeta donde desea guardar los documentos generados"
        If .Show = -1 Then
            saveFolder = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    ' Crea una instancia de la aplicación de Word
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True
    
    ' Abre el documento de Word seleccionado
    Set wordDoc = wordApp.Documents.Open(templatePath)
    
    ' Crea un diccionario de reemplazo para los campos
    Set replaceDic = CreateObject("Scripting.Dictionary")
    
    ' Llena el diccionario de reemplazo con los datos de la tabla de Excel
    rowCount = rng.Rows.Count
    For Each cell In selectedRange.Rows(1).Cells ' Tomamos la primera fila para los nombres de los campos
        replaceDic("«" & cell.Value & "»") = ""
    Next cell
    
    ' Crea una carpeta temporal en la carpeta de archivos temporales del sistema
    tempFolder = Environ("TEMP") & "\tmp-" & Format(Now(), "yyyymmddhhmmss")
    MkDir tempFolder
    
    ' Copia el documento de Word seleccionado a la carpeta temporal
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    ' Genera un archivo de Word por cada registro de la tabla
    For i = 2 To rowCount ' Empezamos desde la segunda fila para los datos reales
        ' Crear un nuevo diccionario para cada fila
        Set replaceDic = CreateObject("Scripting.Dictionary")
        
        ' Llena el diccionario de reemplazo con los datos de la fila actual de la tabla de Excel
        For Each cell In selectedRange.Rows(1).Cells ' Tomamos la primera fila para los nombres de los campos
            replaceDic("«" & cell.Value & "»") = rng.Cells(i, cell.Column).Value
        Next cell
        
        ' Crea una copia del documento de Word en la carpeta temporal
        fs.CopyFile templatePath, tempFolder & "\Documento_" & i & ".docx"
        ' Abre la copia del documento de Word
        Set wordDoc = wordApp.Documents.Open(tempFolder & "\Documento_" & i & ".docx")
        ' Realiza los reemplazos en el documento de Word
        For Each Key In replaceDic.Keys
            Debug.Print CStr(Key)
            If CStr(Key) = "«Descripcion»" Then
                ' Aplicar la función específica para la clave «Descripcion»
                replaceDic(Key) = TransformText(replaceDic(Key))
            End If
            ' Reemplazar en el documento de Word
            WordAppReplaceParagraph wordApp, wordDoc, CStr(Key), CStr(replaceDic(Key))
        Next Key
        FormatRiskLevelCell wordDoc.Tables(1).cell(1, 2)
        ' Guarda y cierra el documento de Word
        ' Antes de guardar el documento de Word
        EliminarUltimasFilasSiEsSalidaPruebaSeguridad wordDoc, replaceDic
        wordDoc.Save
        wordDoc.Close
        
        ' Agregar el documento generado a la lista
        ReDim Preserve documentsList(i - 2)
        documentsList(i - 2) = tempFolder & "\Documento_" & i & ".docx"
    Next i
    
    ' Combina todos los archivos en uno solo
    Dim finalDocumentPath As String
    finalDocumentPath = saveFolder & "\Documento_Consolidado.docx"
    MergeDocuments wordApp, documentsList, finalDocumentPath
    
    ' Mueve la carpeta temporal a la carpeta seleccionada por el usuario
    fs.MoveFolder tempFolder, saveFolder & "\DocumentosGenerados"
    
    ' Cerrar la aplicación de Word
    wordApp.Quit
    Set wordApp = Nothing
    
    ' Muestra un mensaje de éxito
    MsgBox "Se han generado los documentos de Word correctamente.", vbInformation
End Sub





Sub WordAppReplaceParagraph(wordApp As Object, wordDoc As Object, wordToFind As String, replaceWord As String)
    Dim findInRange As Boolean
    
      ' Ir al principio del documento nuevamente
    wordApp.Selection.GoTo What:=1, Which:=1, Name:="1"
    wordApp.ActiveWindow.ActivePane.View.SeekView = 0
    
    ' Bucle para buscar y reemplazar todas las ocurrencias
    Do
        ' Intentar encontrar y reemplazar en el cuerpo del documento
        findInRange = wordApp.Selection.Find.Execute(FindText:=wordToFind)
        
        ' Si se encontró el texto, reemplazarlo
        If findInRange Then
        
       
    
            ' Realizar el reemplazo
            wordApp.Selection.text = replaceWord
            
             ' Ir al principio del documento nuevamente
    wordApp.Selection.GoTo What:=1, Which:=1, Name:="1"
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

Sub EliminarUltimasFilasSiEsSalidaPruebaSeguridad(wordDoc As Object, replaceDic As Object)
    Dim salidaPruebaSeguridadKey As String
    salidaPruebaSeguridadKey = "«SalidaPruebaSeguridad»"
    
    ' Verificar si la clave "«SalidaPruebaSeguridad»" está presente en el diccionario
    If replaceDic.Exists(salidaPruebaSeguridadKey) Then
        ' Verificar si el valor asociado coincide con el texto específico
        If replaceDic(salidaPruebaSeguridadKey) = "La herramienta identificó la vulnerabilidad mediante una prueba específica, ya sea mediante el empleo de una solicitud preparada, la utilización de un plugin o un intento de conexión directa. Esta evaluación confirmó si la respuesta se recibió de manera exitosa. Para acceder a información más detallada sobre la vulnerabilidad, le recomendamos consultar la descripción correspondiente o referirse a la fuente de referencia proporcionada." Then
            ' Eliminar las últimas dos filas de la primera tabla en el documento
            Dim firstTable As Object
            Set firstTable = wordDoc.Tables(1)
            Dim numRows As Integer
            numRows = firstTable.Rows.Count
            
            If numRows > 0 Then
                ' Eliminar la última fila dos veces
                firstTable.Rows(numRows).Delete
                If numRows > 1 Then
                    firstTable.Rows(numRows - 1).Delete
                End If
            End If
        End If
    End If
End Sub
' Crear arvhivo word

Sub MergeDocuments(wordApp As Object, documentsList As Variant, finalDocumentPath As String)
    Dim baseDoc As Object
    Dim sFile As String
    Dim oRng As Object
    Dim i As Integer
    
    On Error GoTo err_Handler
    
    ' Crear un nuevo documento base
    Set baseDoc = wordApp.Documents.Add
    
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
        estilo = r.Range.Cells(1, tbl.ListColumns("Estilo").Index).Value
        seccion = r.Range.Cells(1, tbl.ListColumns("Sección").Index).Value
        descripcion = r.Range.Cells(1, tbl.ListColumns("Descripción").Index).Value
        imagen = r.Range.Cells(1, tbl.ListColumns("Imágenes").Index).Value
        parrafoResultados = r.Range.Cells(1, tbl.ListColumns("Párrafo de resultados").Index).Value
        
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

