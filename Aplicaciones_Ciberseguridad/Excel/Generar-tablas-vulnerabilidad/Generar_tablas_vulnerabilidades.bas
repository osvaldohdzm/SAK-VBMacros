Attribute VB_Name = "Módulo12"
Sub GenerarDocumentosWord()
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
    Dim COLORIKETRIA As String
    
    ' Definir la constante según sea necesario
    COLORIMETRIA = "INAI"  ' Cambiar a "INAI" si es necesario
    
    ' Obtener el texto de la celda y eliminar los caracteres especiales
    cellText = Replace(cell.Range.text, vbCrLf, "")
    cellText = Replace(cellText, vbCr, "")
    cellText = Replace(cellText, vbLf, "")
    cellText = Replace(cellText, Chr(7), "")
    
    ' Realizar la comparación utilizando el texto de la celda sin caracteres especiales
    If COLORIMETRIA = "BANOBRAS" Then
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
    ElseIf COLORIMETRIA = "INAI" Then
        ' Asignar colores para INAI según especificación
        Select Case cellText
            Case "CRÍTICA"
                cell.Shading.BackgroundPatternColor = RGB(255, 0, 0) ' Rojo para "CRÍTICA"
                cell.Range.Font.Color = RGB(255, 255, 255) ' Texto blanco para "CRÍTICA"
            Case "ALTA"
                cell.Shading.BackgroundPatternColor = RGB(255, 102, 0) ' Naranja para "ALTA"
                cell.Range.Font.Color = RGB(255, 255, 255) ' Texto blanco para "ALTA"
            Case "MEDIA"
                cell.Shading.BackgroundPatternColor = RGB(255, 192, 0) ' Amarillo para "MEDIA"
                cell.Range.Font.Color = RGB(0, 0, 0) ' Texto negro para "MEDIA"
            Case "BAJA"
                cell.Shading.BackgroundPatternColor = RGB(0, 176, 80) ' Verde para "BAJA"
                cell.Range.Font.Color = RGB(255, 255, 255) ' Texto blanco para "BAJA"
        End Select
    End If
End Sub

Function TransformText(text As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' Configurar la expresión regular para encontrar saltos de línea o saltos de carro sin un punto antes y no seguidos de paréntesis ni de guión
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = "([^.()\r\n-])(?![^(]*\)|[-])[^\S\r\n]*[\r\n]+" ' Expresión regular para encontrar saltos de línea o saltos de carro sin un punto antes y no seguidos de paréntesis ni de guión
    End With
    
    ' Realizar la transformación: quitar caracteres especiales y aplicar la expresión regular
    TransformText = regEx.Replace(Replace(text, Chr(7), ""), "$1 ")
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

