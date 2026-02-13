Attribute VB_Name = "Módulo11"
Sub GEN_001_MarkTables()
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

Sub BorrarContenidoCelda()
    Dim celdaActual As cell
    Dim rango As Range

    ' 1. Verificar si el cursor está dentro de una tabla
    If Selection.Information(wdWithInTable) = False Then
        MsgBox "No estás dentro de una tabla. Por favor haz clic en una celda.", vbExclamation
        Exit Sub
    End If

    ' 2. Capturar la celda actual (donde está el cursor)
    Set celdaActual = Selection.Cells(1)

    ' 3. Definir el rango del contenido
    Set rango = celdaActual.Range

    ' IMPORTANTE: Word incluye una "marca de fin de celda" invisible al final.
    ' Si borras esa marca, la celda se rompe o se fusiona.
    ' Retrocedemos 1 carácter para borrar solo el contenido.
    rango.End = rango.End - 1

    ' 4. Borrar todo (texto, imágenes y tablas anidadas dentro)
    rango.Delete
End Sub

Sub GEN_002_MarkInlineCharts()
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

Sub GEN_003_FormatearTabla()
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


Sub GEN_004_ActualizarCamposSEQ()
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

Sub GEN_005_CambiarMontserrat10a11()
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

Sub GEN_006_FormatoNegritaViñetas()
    Dim p As Paragraph
    Dim strTexto As String
    Dim posDosPuntos As Integer
    Dim rng As Range

    ' Recorrer cada párrafo en la selección
    For Each p In Selection.Paragraphs
        ' Asignar el texto seleccionado a una cadena
        strTexto = p.Range.Text
        
        ' Buscar la posición de los dos puntos
        posDosPuntos = InStr(strTexto, ":")

        ' Si se encuentran los dos puntos
        If posDosPuntos > 0 Then
            ' Seleccionar el rango desde el inicio hasta los dos puntos
            Set rng = p.Range
            rng.Start = p.Range.Start
            rng.End = p.Range.Start + posDosPuntos - 1
            rng.Font.Bold = True ' Aplicar negrita

            ' Seleccionar el rango después de los dos puntos y quitar la negrita
            Set rng = p.Range
            rng.Start = p.Range.Start + posDosPuntos
            rng.End = p.Range.End
            rng.Font.Bold = False ' Quitar negrita
        End If
    Next p
End Sub


Sub GEN_007_FormatoNegritaViñetasGuion()
    Dim p As Paragraph
    Dim strTexto As String
    Dim posDosPuntos As Integer
    Dim rng As Range
    
    ' Recorrer cada párrafo en la selección
    For Each p In Selection.Paragraphs
        ' Asignar el texto del párrafo a una cadena
        strTexto = p.Range.Text
        
        ' Verificar si el párrafo comienza con "- "
        If Left(Trim(strTexto), 2) = "- " Then
            ' Buscar la posición de los dos puntos
            posDosPuntos = InStr(strTexto, ":")
            
            ' Si se encuentran los dos puntos
            If posDosPuntos > 0 Then
                ' Seleccionar el rango desde el inicio hasta los dos puntos
                Set rng = p.Range
                rng.Start = p.Range.Start
                rng.End = p.Range.Start + posDosPuntos - 1
                rng.Font.Bold = True ' Aplicar negrita
                
                ' Seleccionar el rango después de los dos puntos y quitar la negrita
                Set rng = p.Range
                rng.Start = p.Range.Start + posDosPuntos
                rng.End = p.Range.End
                rng.Font.Bold = False ' Quitar negrita
            End If
        End If
    Next p
End Sub



Sub GEN_008_BuscarPalabrasYGenerarCSV()
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

Sub GEN_009_ForzarRenumerarCaptionImagenes()

    ' Definir variables
    Dim rng As Range
    Dim contador As Integer
    Dim encontrado As Boolean
    
    ' Inicializar contador
    contador = 1

    ' Establecer el rango para buscar en todo el documento
    Set rng = ActiveDocument.Content
    
    ' Inicializar la búsqueda
    With rng.Find
        .Text = "Imagen [0-9]{1,}:"
        .MatchWildcards = True ' Usar comodines para encontrar números
        .Forward = True
        
        ' Bucle para encontrar todas las ocurrencias
        Do While .Execute
            ' Verificar si se encontró el patrón
            If .Found Then
                ' Ajustar el texto encontrado con el nuevo número
                rng.Text = "Imagen " & contador & ":"
                contador = contador + 1 ' Incrementar el contador
            End If
        Loop
    End With
    
    ' Notificar al usuario que se completó la renumeración
    MsgBox "Renumeración de imágenes completada.", vbInformation

End Sub

Sub GEN_010_ReplacePicoParentesis()
    Dim doc As Document
    Dim rng As Range
    Dim foundText As String
    Dim newValue As String
    Dim inputDict As Object
    Dim headerFooter As headerFooter
    Dim shape As shape

    ' Inicializa el objeto para almacenar los valores ya reemplazados
    Set inputDict = CreateObject("Scripting.Dictionary")

    ' Establece el documento actual
    Set doc = ActiveDocument

    ' Recorre todo el contenido del documento, incluidos los encabezados y pies de página
    Set rng = doc.Content

    ' Define el patrón de búsqueda para las cadenas con pico paréntesis
    With rng.Find
        .Text = "«*»"
        .MatchWildcards = True
        .Forward = True

        ' Realiza la búsqueda en todo el documento
        Do While .Execute
            foundText = rng.Text

            ' Verifica si ya se solicitó el valor para esta cadena
            If Not inputDict.Exists(foundText) Then
                ' Solicita al usuario el nuevo valor
                newValue = InputBox("Ingresa el nuevo valor para " & foundText, "Reemplazar texto")
                
                ' Almacena el valor ingresado o conserva el valor actual si está vacío
                If newValue <> "" Then
                    inputDict.Add foundText, newValue
                Else
                    inputDict.Add foundText, foundText
                End If
            End If

            ' Reemplaza el texto con el valor ingresado o mantiene el original
            If inputDict.Exists(foundText) Then
                rng.Text = inputDict(foundText)
            End If

            ' Actualiza el rango para continuar la búsqueda
            rng.Collapse wdCollapseEnd
        Loop
    End With

    ' Recorre los encabezados y pies de página
    For Each headerFooter In doc.Sections(1).Headers
        Set rng = headerFooter.Range
        Call GEN_011_ReplaceInRange(rng, inputDict)
    Next headerFooter
    
    For Each headerFooter In doc.Sections(1).Footers
        Set rng = headerFooter.Range
        Call GEN_011_ReplaceInRange(rng, inputDict)
    Next headerFooter
    
    ' Recorre las autormas en el documento (Shapes tipo msoTextBox)
    For Each shape In doc.Shapes
        If shape.Type = msoTextBox Then
            ' Llama a la función ReplaceInRange para reemplazar texto dentro de la autorma
            Call GEN_011_ReplaceInRange(shape.TextFrame.TextRange, inputDict)
        End If
    Next shape

    ' Limpia los objetos
    Set rng = Nothing
    Set inputDict = Nothing
End Sub

Sub GEN_011_ReplaceInRange(rng As Range, inputDict As Object)
    Dim foundText As String
    Dim newValue As String
    
    ' Define el patrón de búsqueda para las cadenas con pico paréntesis
    With rng.Find
        .Text = "«*»"
        .MatchWildcards = True
        .Forward = True
        
        ' Realiza la búsqueda en el rango especificado
        Do While .Execute
            foundText = rng.Text

            ' Verifica si ya se solicitó el valor para esta cadena
            If Not inputDict.Exists(foundText) Then
                ' Solicita al usuario el nuevo valor
                newValue = InputBox("Ingresa el nuevo valor para " & foundText, "Reemplazar texto")
                
                ' Almacena el valor ingresado o conserva el valor actual si está vacío
                If newValue <> "" Then
                    inputDict.Add foundText, newValue
                Else
                    inputDict.Add foundText, foundText
                End If
            End If

            ' Reemplaza el texto con el valor ingresado o mantiene el original
            If inputDict.Exists(foundText) Then
                rng.Text = inputDict(foundText)
            End If

            ' Actualiza el rango para continuar la búsqueda
            rng.Collapse wdCollapseEnd
        Loop
    End With
End Sub

Sub GEN_012_ModificarEstiloTitulo3()
    Dim est As Style
    
    ' Acceder al estilo Título 3 del documento activo
    Set est = ActiveDocument.Styles(wdStyleHeading3)
    
    With est
        ' 1. Configurar la Fuente
        .Font.Name = "Roboto Slap" ' Puedes cambiar la tipografía aquí
        .Font.Size = 12        ' Tamaño de letra
        .Font.Bold = True      ' Negrita
        .Font.ColorIndex = wdWhite ' Letra Blanca
        
        ' 2. Configurar el Párrafo y el Fondo (Sombreado)
        With .ParagraphFormat
            .Alignment = wdAlignParagraphLeft ' Alineación a la izquierda
            .SpaceBefore = 12  ' Espacio antes del título
            .SpaceAfter = 6    ' Espacio después del título
            
            ' Aplicar el fondo azul sólido
            .Shading.BackgroundPatternColor = wdColorBlue
        End With
        
        ' 3. Asegurar que no tenga subrayado de línea
        .Font.Underline = wdUnderlineNone
    End With

    MsgBox "El estilo 'Título 3' ha sido actualizado con éxito."
End Sub

