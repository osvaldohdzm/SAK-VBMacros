Attribute VB_Name = "M�dulo11"
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
            ' Revisa si el texto tiene la fuente Montserrat y tama�o 10
            If rng.Font.Name = "Montserrat" And rng.Font.Size = 10 Then
                ' Cambia el tama�o de fuente a 11
                rng.Font.Size = 11
            End If
            ' Mueve al siguiente rango
            Set rng = rng.NextStoryRange
        Loop Until rng Is Nothing
    Next i
    
    ' Mensaje de finalizaci�n
    MsgBox "Cambio completado de Montserrat 10 a 11", vbInformation
End Sub

Sub GEN006_FormatoNegritaVi�etas()
    Dim p As Paragraph
    Dim strTexto As String
    Dim posDosPuntos As Integer
    Dim rng As Range

    ' Recorrer cada p�rrafo en la selecci�n
    For Each p In Selection.Paragraphs
        ' Asignar el texto seleccionado a una cadena
        strTexto = p.Range.Text
        
        ' Buscar la posici�n de los dos puntos
        posDosPuntos = InStr(strTexto, ":")

        ' Si se encuentran los dos puntos
        If posDosPuntos > 0 Then
            ' Seleccionar el rango desde el inicio hasta los dos puntos
            Set rng = p.Range
            rng.Start = p.Range.Start
            rng.End = p.Range.Start + posDosPuntos - 1
            rng.Font.Bold = True ' Aplicar negrita

            ' Seleccionar el rango despu�s de los dos puntos y quitar la negrita
            Set rng = p.Range
            rng.Start = p.Range.Start + posDosPuntos
            rng.End = p.Range.End
            rng.Font.Bold = False ' Quitar negrita
        End If
    Next p
End Sub


Sub GEN007_FormatoNegritaVi�etasGuion()
    Dim p As Paragraph
    Dim strTexto As String
    Dim posDosPuntos As Integer
    Dim rng As Range
    
    ' Recorrer cada p�rrafo en la selecci�n
    For Each p In Selection.Paragraphs
        ' Asignar el texto del p�rrafo a una cadena
        strTexto = p.Range.Text
        
        ' Verificar si el p�rrafo comienza con "- "
        If Left(Trim(strTexto), 2) = "- " Then
            ' Buscar la posici�n de los dos puntos
            posDosPuntos = InStr(strTexto, ":")
            
            ' Si se encuentran los dos puntos
            If posDosPuntos > 0 Then
                ' Seleccionar el rango desde el inicio hasta los dos puntos
                Set rng = p.Range
                rng.Start = p.Range.Start
                rng.End = p.Range.Start + posDosPuntos - 1
                rng.Font.Bold = True ' Aplicar negrita
                
                ' Seleccionar el rango despu�s de los dos puntos y quitar la negrita
                Set rng = p.Range
                rng.Start = p.Range.Start + posDosPuntos
                rng.End = p.Range.End
                rng.Font.Bold = False ' Quitar negrita
            End If
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
        MsgBox "No se seleccion� ning�n archivo.", vbExclamation
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
    
    ' Procesar cada l�nea del archivo TXT
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

Sub ForzarRenumerarCaptionImagenes()

    ' Definir variables
    Dim rng As Range
    Dim contador As Integer
    Dim encontrado As Boolean
    
    ' Inicializar contador
    contador = 1

    ' Establecer el rango para buscar en todo el documento
    Set rng = ActiveDocument.Content
    
    ' Inicializar la b�squeda
    With rng.Find
        .Text = "Imagen [0-9]{1,}:"
        .MatchWildcards = True ' Usar comodines para encontrar n�meros
        .Forward = True
        
        ' Bucle para encontrar todas las ocurrencias
        Do While .Execute
            ' Verificar si se encontr� el patr�n
            If .Found Then
                ' Ajustar el texto encontrado con el nuevo n�mero
                rng.Text = "Imagen " & contador & ":"
                contador = contador + 1 ' Incrementar el contador
            End If
        Loop
    End With
    
    ' Notificar al usuario que se complet� la renumeraci�n
    MsgBox "Renumeraci�n de im�genes completada.", vbInformation

End Sub

Sub GEN_005_ReplacePicoParentesis()
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

    ' Recorre todo el contenido del documento, incluidos los encabezados y pies de p�gina
    Set rng = doc.Content

    ' Define el patr�n de b�squeda para las cadenas con pico par�ntesis
    With rng.Find
        .Text = "�*�"
        .MatchWildcards = True
        .Forward = True

        ' Realiza la b�squeda en todo el documento
        Do While .Execute
            foundText = rng.Text

            ' Verifica si ya se solicit� el valor para esta cadena
            If Not inputDict.Exists(foundText) Then
                ' Solicita al usuario el nuevo valor
                newValue = InputBox("Ingresa el nuevo valor para " & foundText, "Reemplazar texto")
                
                ' Almacena el valor ingresado o conserva el valor actual si est� vac�o
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

            ' Actualiza el rango para continuar la b�squeda
            rng.Collapse wdCollapseEnd
        Loop
    End With

    ' Recorre los encabezados y pies de p�gina
    For Each headerFooter In doc.Sections(1).Headers
        Set rng = headerFooter.Range
        Call ReplaceInRange(rng, inputDict)
    Next headerFooter
    
    For Each headerFooter In doc.Sections(1).Footers
        Set rng = headerFooter.Range
        Call ReplaceInRange(rng, inputDict)
    Next headerFooter
    
    ' Recorre las autormas en el documento (Shapes tipo msoTextBox)
    For Each shape In doc.Shapes
        If shape.Type = msoTextBox Then
            ' Llama a la funci�n ReplaceInRange para reemplazar texto dentro de la autorma
            Call ReplaceInRange(shape.TextFrame.TextRange, inputDict)
        End If
    Next shape

    ' Limpia los objetos
    Set rng = Nothing
    Set inputDict = Nothing
End Sub

Sub ReplaceInRange(rng As Range, inputDict As Object)
    Dim foundText As String
    Dim newValue As String
    
    ' Define el patr�n de b�squeda para las cadenas con pico par�ntesis
    With rng.Find
        .Text = "�*�"
        .MatchWildcards = True
        .Forward = True
        
        ' Realiza la b�squeda en el rango especificado
        Do While .Execute
            foundText = rng.Text

            ' Verifica si ya se solicit� el valor para esta cadena
            If Not inputDict.Exists(foundText) Then
                ' Solicita al usuario el nuevo valor
                newValue = InputBox("Ingresa el nuevo valor para " & foundText, "Reemplazar texto")
                
                ' Almacena el valor ingresado o conserva el valor actual si est� vac�o
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

            ' Actualiza el rango para continuar la b�squeda
            rng.Collapse wdCollapseEnd
        Loop
    End With
End Sub

