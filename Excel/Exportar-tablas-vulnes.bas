Attribute VB_Name = "Módulo1"
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
    For Each cell In selectedRange.Rows(1).Cells
        replaceDic("«" & cell.Value & "»") = ""
        For i = 1 To rowCount
            replaceDic("«" & cell.Value & "»") = replaceDic("«" & cell.Value & "»") & vbCrLf & rng.Cells(i, cell.Column).Value
        Next i
    Next cell
    
    ' Crea una carpeta temporal en la carpeta de archivos temporales del sistema
    tempFolder = Environ("TEMP") & "\tmp-" & Format(Now(), "yyyymmddhhmmss")
    MkDir tempFolder
    
    ' Copia el documento de Word seleccionado a la carpeta temporal
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.CopyFile templatePath, tempFolder & "\Plantilla.docx"
    
    ' Genera un archivo de Word por cada registro de la tabla
    For i = 1 To rowCount
        ' Crear un nuevo diccionario para cada fila
        Set replaceDic = CreateObject("Scripting.Dictionary")
        
        ' Llena el diccionario de reemplazo con los datos de la fila actual de la tabla de Excel
        For Each cell In selectedRange.Rows(1).Cells
            replaceDic("«" & cell.Value & "»") = rng.Cells(i, cell.Column).Value
        Next cell
        
        ' Crea una copia del documento de Word en la carpeta temporal
        fs.CopyFile tempFolder & "\Plantilla.docx", tempFolder & "\Documento_" & i & ".docx"
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
    Next i
    
    ' Cierra la aplicación de Word
    wordApp.Quit
    
    ' Combina todos los archivos en uno solo
    Dim outputFile As Object
    Set outputFile = fs.CreateTextFile(tempFolder & "\Documento_Consolidado.docx")
    For i = 1 To rowCount
        Dim content As String
        content = fs.OpenTextFile(tempFolder & "\Documento_" & i & ".docx").ReadAll
        outputFile.WriteLine content
        fs.DeleteFile tempFolder & "\Documento_" & i & ".docx"
    Next i
    outputFile.Close
    
    ' Mueve la carpeta temporal y el archivo consolidado a la carpeta seleccionada por el usuario
    fs.MoveFolder tempFolder, saveFolder & "\DocumentosGenerados"
    
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
        Case "CRÍTICO"
            cell.Shading.BackgroundPatternColor = 10498160
            cell.Range.Font.Color = 16777215
        Case "ALTO"
            cell.Shading.BackgroundPatternColor = 255
            cell.Range.Font.Color = 16777215
        Case "MEDIO"
            cell.Shading.BackgroundPatternColor = 65535
            cell.Range.Font.Color = 0
        Case "BAJO"
            cell.Shading.BackgroundPatternColor = 5287936
            cell.Range.Font.Color = 16777215
    End Select
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

