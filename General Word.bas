Attribute VB_Name = "Módulo11"
Sub MarkTables()
    Dim doc As Document
    Dim table As table
    Dim count As Integer
    count = 1
    
    Set doc = ActiveDocument ' Documento activo
    
    For Each table In doc.Tables
        On Error Resume Next ' Ignorar errores para tablas sin celdas
        table.cell(1, 1).Range.Text = "Table " & CStr(count)
        On Error GoTo 0 ' Restaurar manejo normal de errores
        count = count + 1
    Next table
End Sub


Sub MarkInlineCharts()
    Dim doc As Document
    Dim inlineShape As inlineShape
    Dim count As Integer
    count = 1
    
    Set doc = ActiveDocument ' Documento activo
    
    For Each inlineShape In doc.InlineShapes
        On Error Resume Next ' Ignorar errores para elementos sin texto alternativo
        inlineShape.AlternativeText = "MyGrafico " & CStr(count)
        On Error GoTo 0 ' Restaurar manejo normal de errores
        count = count + 1
    Next inlineShape
End Sub

Sub FormatearTabla()
    Dim tbl As table
    Dim row As row
    Dim isFirstRow As Boolean
    
    ' Comprobar si hay al menos una tabla seleccionada
    If Selection.Tables.count = 0 Then
        MsgBox "No hay ninguna tabla seleccionada.", vbExclamation
        Exit Sub
    End If
    
    ' Iterar sobre cada tabla seleccionada
    For Each tbl In Selection.Tables
        ' Formatear la tabla con borde de 1/2 punto y color RGB(0, 112, 192)
        With tbl
            .Borders.Enable = True
            .Borders.InsideLineStyle = wdLineStyleSingle
            .Borders.OutsideLineStyle = wdLineStyleSingle
            .Borders.OutsideLineWidth = wdLineWidth050pt
            .Borders.InsideColor = RGB(0, 112, 192)
            .Borders.OutsideColor = RGB(0, 112, 192)
        End With
        
        ' Formatear la primera fila (cabecera)
        Set row = tbl.Rows(1)
        row.Range.Shading.BackgroundPatternColor = RGB(0, 112, 192)
        row.Range.Font.Color = RGB(255, 255, 255)
        row.Range.ParagraphFormat.SpaceBefore = 0
        row.Range.ParagraphFormat.SpaceAfter = 0
        row.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter 'Centrar texto
        
        ' Iterar sobre las filas restantes
        isFirstRow = True
        For Each row In tbl.Rows
            If Not isFirstRow Then
                ' Aplicar formato a filas no cabecera
                Dim cell As cell
                For Each cell In row.Cells
                    ' Verificar si la celda tiene un color de fondo predeterminado (blanco)
                    If cell.Range.Shading.BackgroundPatternColorIndex = wdColorAutomatic Then
                        If row.Index Mod 2 = 0 Then
                            ' Filas pares
                            row.Range.Shading.BackgroundPatternColor = RGB(255, 255, 255)
                            row.Range.Font.Color = RGB(0, 0, 0) ' Color de letra negro
                        Else
                            ' Filas impares
                            row.Range.Shading.BackgroundPatternColor = RGB(217, 226, 243) ' Color de fondo blanco
                            row.Range.Font.Color = RGB(0, 0, 0) ' Color de letra negro
                        End If
                    End If
                Next cell
                row.Range.ParagraphFormat.SpaceBefore = 0
                row.Range.ParagraphFormat.SpaceAfter = 0
                row.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter 'Centrar texto
                
                ' Centrar contenido verticalmente
                row.Cells.VerticalAlignment = wdCellAlignVerticalCenter
            End If
            isFirstRow = False
        Next row
    Next tbl
    
    MsgBox "La tabla seleccionada ha sido formateada correctamente.", vbInformation
End Sub


Sub ActualizarCamposSEQ()
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Recorre todos los campos del documento
    For Each campo In doc.Fields
        ' Comprueba si el campo es de tipo SEQ
        If campo.Type = wdFieldSequence Then
            ' Actualiza el campo
            campo.Update
        End If
    Next campo
End Sub

Sub CambiarMontserrat10a11()
    Dim rng As Range
    Dim doc As Document
    Dim i As Integer
    
    ' Establece el documento activo
    Set doc = ActiveDocument
    
    ' Itera sobre todos los rangos del documento
    For i = 1 To doc.StoryRanges.count
        Set rng = doc.StoryRanges(i)
        Do
            ' Revisa si el texto tiene la fuente Montserrat y tamaño 10
            If rng.Font.Name = "Montserrat" And rng.Font.Size = 10 Then
                ' Cambia el tamaño de fuente a 11
                rng.Font.Size = 11
            End If
            ' Mueve al siguiente rango
            Set rng = rng.NextStoryRange
        Loop Until rng Is Nothing
    Next i
    
    ' Mensaje de finalización
    MsgBox "Cambio completado de Montserrat 10 a 11", vbInformation
End Sub

Sub FormatoNegritaViñetas()
    Dim p As Paragraph
    Dim strTexto As String
    Dim arrLineas As Variant
    Dim i As Integer
    Dim posDosPuntos As Integer
    Dim rng As Range
    
    ' Recorrer cada párrafo en la selección
    For Each p In Selection.Paragraphs
        If p.Range.ListFormat.ListType = wdListBullet Then ' Verificar que sea una viñeta
        
            ' Asignar el texto seleccionado a una cadena
            strTexto = p.Range.Text
            
            ' Dividir el texto en líneas
            arrLineas = Split(strTexto, vbCrLf)
            
            ' Recorrer cada línea
            For i = LBound(arrLineas) To UBound(arrLineas)
                ' Obtener la posición de los dos puntos
                posDosPuntos = InStr(arrLineas(i), ":")
                
                ' Seleccionar el texto desde el primer carácter hasta posDosPuntos y aplicar negrita
                Set rng = p.Range
                rng.MoveStart unit:=wdCharacter, count:=0
                rng.MoveEnd unit:=wdCharacter, count:=posDosPuntos - 1
                rng.Font.Bold = True
                
                ' Seleccionar el texto desde posDosPuntos hasta el final y quitar negrita
                rng.MoveStart unit:=wdCharacter, count:=posDosPuntos - 1
                rng.MoveEnd unit:=wdCharacter, count:=Len(arrLineas(i)) - posDosPuntos + 1
                rng.Font.Bold = False
            Next i
            
        End If
    Next p
End Sub


Sub BuscarPalabrasYGenerarCSV()
    Dim dlgOpen As FileDialog
    Dim archivoTXT As String
    Dim archivoCSV As String
    Dim fileTXT As Integer
    Dim fileCSV As Integer
    Dim palabra As String
    Dim textoDocumento As String
    Dim encontrado As Boolean

    ' Solicitar el archivo .txt
    Set dlgOpen = Application.FileDialog(msoFileDialogOpen)
    dlgOpen.Title = "Seleccionar archivo TXT"
    dlgOpen.Filters.Clear
    dlgOpen.Filters.Add "Text Files", "*.txt"
    
    If dlgOpen.Show = -1 Then
        archivoTXT = dlgOpen.SelectedItems(1)
    Else
        MsgBox "No se seleccionó ningún archivo.", vbExclamation
        Exit Sub
    End If

    ' Configurar archivo CSV de salida
    archivoCSV = Left(archivoTXT, InStrRev(archivoTXT, ".")) & "csv"
    
    ' Abrir archivo TXT para lectura
    fileTXT = FreeFile
    Open archivoTXT For Input As fileTXT
    
    ' Abrir archivo CSV para escritura
    fileCSV = FreeFile
    Open archivoCSV For Output As fileCSV
    
    ' Leer todo el contenido del documento de Word
    textoDocumento = ActiveDocument.Content.Text
    
    ' Procesar cada línea del archivo TXT
    Do While Not EOF(fileTXT)
        Line Input #fileTXT, palabra
        encontrado = InStr(1, textoDocumento, palabra, vbTextCompare) > 0
        
        ' Escribir resultado en el archivo CSV
        If encontrado Then
            Print #fileCSV, palabra & ",FOUND"
        Else
            Print #fileCSV, palabra & ",NOT FOUND"
        End If
    Loop
    
    ' Cerrar archivos
    Close fileTXT
    Close fileCSV
    
    MsgBox "Proceso completado. Archivo CSV generado: " & archivoCSV, vbInformation
End Sub

