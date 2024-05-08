Attribute VB_Name = "NewMacros"
Sub CopiarImagenesAPowerPoint()
    Dim PowerPointApp As Object
    Dim PowerPointPresentation As Object
    Dim PowerPointSlide As Object
    Dim Imagen As InlineShape
    Dim PlantillaPowerPoint As String
    Dim NuevaPresentacion As String
    Dim i As Integer
    Dim textoDespuesImagenes As String
    
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
            MsgBox "No se seleccion� ninguna plantilla. El proceso se cancel�.", vbExclamation
            Exit Sub
        End If
    End With
    
    ' Crear una nueva presentaci�n en PowerPoint basada en la plantilla
    NuevaPresentacion = rutaDocumento & "\Presentacion_avanzada.pptx"
    FileCopy PlantillaPowerPoint, NuevaPresentacion
    Set PowerPointApp = CreateObject("PowerPoint.Application")
    PowerPointApp.Visible = True
    Set PowerPointPresentation = PowerPointApp.Presentations.Open(NuevaPresentacion)
    
    ' Buscar la diapositiva cuyo t�tulo es "Plantilla de captura" y duplicarla
    For Each PowerPointSlide In PowerPointPresentation.Slides
        If PowerPointSlide.Shapes.HasTitle Then
            If PowerPointSlide.Shapes.Title.TextFrame.TextRange.Text = "Plantilla de captura" Then
                Exit For
            End If
        End If
    Next PowerPointSlide
    
    ' Verificar si se encontr� la diapositiva "Plantilla de captura"
    If PowerPointSlide Is Nothing Then
        MsgBox "No se encontr� la diapositiva con el t�tulo 'Plantilla de captura'. El proceso se cancel�.", vbExclamation
        PowerPointPresentation.Close
        PowerPointApp.Quit
        Exit Sub
    End If
    
    ' Copiar cada imagen del documento de Word y pegarla en una nueva diapositiva basada en la diapositiva de plantilla duplicada
    For Each Imagen In ActiveDocument.InlineShapes
        ' Verificar si la forma es una imagen
        If Imagen.Type = wdInlineShapePicture Then
            ' Duplicar la diapositiva de plantilla despu�s de la original
            Set newSlide = PowerPointSlide.Duplicate
            newSlide.MoveTo PowerPointPresentation.Slides.Count ' Mover la diapositiva duplicada al final
            
            ' Pegar la imagen en la nueva diapositiva
            Imagen.Select
            Selection.Copy
            newSlide.Shapes.PasteSpecial DataType:=ppPasteShape
            
            ' Obtener el �ndice de la nueva forma en la diapositiva
            Dim NewShapeIndex As Integer
            NewShapeIndex = newSlide.Shapes.Count

            With newSlide.Shapes(NewShapeIndex)
                ' Verificar si el ancho es mayor a 16 y ajustarlo si es necesario
                If .Width > CentimetersToPoints(16) Then
                    .Width = CentimetersToPoints(16)
                End If
                
                ' Verificar si el ancho es menor a 16 y ajustar el alto si es necesario
                If .Width < CentimetersToPoints(16) Then
                    .Height = CentimetersToPoints(10)
                End If
                
                ' Centrar la imagen horizontalmente
                .Left = (PowerPointPresentation.PageSetup.SlideWidth - .Width) / 2
                
                ' Centrar la imagen verticalmente
                .Top = (PowerPointPresentation.PageSetup.SlideHeight - .Height) / 2
            End With
            
            ' Obtener el p�rrafo inmediatamente despu�s de la imagen en Word
            Dim nextParagraph As Paragraph
            Set nextParagraph = Imagen.Range.Paragraphs(1).Next
            
            ' Obtener el texto del p�rrafo
            textoDespuesImagenes = Trim(nextParagraph.Range.Text)
            
            ' Verificar y eliminar el �ltimo car�cter si es un salto de l�nea
            If Right(textoDespuesImagenes, 1) = vbCr Then
                textoDespuesImagenes = Left(textoDespuesImagenes, Len(textoDespuesImagenes) - 1)
            End If
            
            ' Ajustar el t�tulo de la nueva diapositiva
            newSlide.Shapes.Title.TextFrame.TextRange.Text = textoDespuesImagenes
        End If
    Next Imagen
    
    ' Guardar y cerrar la presentaci�n de PowerPoint
    PowerPointPresentation.Save
    PowerPointPresentation.Close
    Set PowerPointPresentation = Nothing
    PowerPointApp.Quit
    
    ' Limpiar objetos
    Set PowerPointApp = Nothing
    
    MsgBox "Las im�genes se han copiado a la presentaci�n de PowerPoint.", vbInformation, "Proceso completado"
End Sub

