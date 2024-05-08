Attribute VB_Name = "NewMacros"
Sub CopiarImagenesAPowerPoint()
    Dim PowerPointApp As Object
    Dim PowerPointPresentation As Object
    Dim PowerPointSlide As Object
    Dim Imagen As InlineShape
    Dim PlantillaPowerPoint As String
    Dim NuevaPresentacion As String
    Dim titulo As String
    Dim textoDespuesImagenes As String
    Dim PowerPointTitleSlide As Object ' Variable para almacenar la diapositiva de título actual
    
    ' Guardar la ruta del documento de Word actual
    Dim rutaDocumento As String
    rutaDocumento = ActiveDocument.Path
    
    ' Seleccionar la plantilla de PowerPoint
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Seleccionar Plantilla de PowerPoint"
        .Filters.Clear
        .Filters.Add "Plantillas de PowerPoint", "*.pptx"
        
        If .Show = -1 Then
            PlantillaPowerPoint = .SelectedItems(1)
        Else
            MsgBox "No se seleccionó ninguna plantilla. El proceso se canceló.", vbExclamation
            Exit Sub
        End If
    End With
    
    ' Crear una nueva presentación en PowerPoint basada en la plantilla
    NuevaPresentacion = rutaDocumento & "\Presentacion_avanzada.pptx"
    FileCopy PlantillaPowerPoint, NuevaPresentacion
    Set PowerPointApp = CreateObject("PowerPoint.Application")
    PowerPointApp.Visible = True
    Set PowerPointPresentation = PowerPointApp.Presentations.Open(NuevaPresentacion)
    
    ' Recorrer el documento para buscar los títulos y crear diapositivas
    For Each parrafo In ActiveDocument.Paragraphs
        titulo = Trim(parrafo.Range.Text)
        
        ' Verificar si el párrafo tiene un estilo de título específico
        If parrafo.Style = "Título 1" Or parrafo.Style = "Título 2" Or parrafo.Style = "Título 3" Then
            ' Crear una nueva diapositiva con el título del párrafo
            Set PowerPointTitleSlide = PowerPointPresentation.Slides.Add(PowerPointPresentation.Slides.Count + 1, 1) ' Usamos ppLayoutTitle para agregar diapositivas con título
            PowerPointTitleSlide.Shapes(1).TextFrame.TextRange.Text = titulo
        ElseIf parrafo.Range.InlineShapes.Count > 0 Then
            ' Verificar si el párrafo tiene una imagen
            For Each Imagen In parrafo.Range.InlineShapes
                ' Copiar la imagen y pegarla en la diapositiva justo después de la diapositiva de título actual
                Imagen.Select
                Selection.Copy
                Set PowerPointSlide = PowerPointPresentation.Slides.Add(PowerPointTitleSlide.SlideIndex + 1, 11)
                PowerPointSlide.Shapes.PasteSpecial DataType:=ppPasteEnhancedMetafile
            Next Imagen
        End If
    Next parrafo
    
    ' Guardar y cerrar la presentación de PowerPoint
    PowerPointPresentation.Save
    PowerPointPresentation.Close
    Set PowerPointPresentation = Nothing
    PowerPointApp.Quit
    
    ' Limpiar objetos
    Set PowerPointApp = Nothing
    
    MsgBox "Las imágenes y títulos se han copiado a la presentación de PowerPoint.", vbInformation, "Proceso completado"
End Sub

