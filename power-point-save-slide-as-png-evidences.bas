Attribute VB_Name = "Módulo1"
Sub SaveSlidesAsPNG()
    Dim sld As Slide
    Dim path As String
    Dim fileName As String
    Dim i As Integer
    
    path = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
    
    ' Recorre todas las diapositivas en la presentación
    For Each sld In ActivePresentation.Slides
        ' Incrementa el contador
        i = i + 1
        
        ' Formatea el número de la evidencia con dos dígitos, como "01", "02", etc.
        fileName = "Evidencia-" & Format(i, "00") & ".png"
        
        ' Exporta la diapositiva como PNG
        sld.Export path & fileName, "PNG"
    Next sld
End Sub
