Attribute VB_Name = "Módulo1"
Sub CopyCurrentSlideAsEnhancedMetafile()
    ' Copiar la diapositiva actual
    ActiveWindow.View.Slide.Copy

    ' Pegar la diapositiva en el portapapeles como un Metafile mejorado (mayor calidad)
    ActiveWindow.View.PasteSpecial DataType:=ppPasteEnhancedMetafile
End Sub

