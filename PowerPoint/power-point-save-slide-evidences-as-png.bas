Attribute VB_Name = "Módulo1"
Sub SaveSlidesEvidencesAsPNG()
    Dim sld As Slide
    Dim path As String
    Dim fileName As String
    Dim currentDate As String
    Dim sPath As String
    Dim i As Integer
    
    ' Obtener la fecha y hora actual para el nombre del archivo
    currentDate = Format(Now(), "yyyy-mm-dd-hh-nn")

    ' Obtener la ubicación de la carpeta "Evidencias" en la ubicación de la presentación
    sPath = ActivePresentation.path
    path = sPath & "\Evidencias\"
    
    ' Comprobar si la carpeta "Evidencias" existe, si no, crearla
    If Dir(path, vbDirectory) = "" Then
        MkDir path
    End If
    
    ' Inicializar el contador
    i = 1
    
    ' Recorrer todas las diapositivas en la presentación
    For Each sld In ActivePresentation.Slides
        ' Formar el nombre del archivo utilizando el contador y la fecha y hora actual
        fileName = "Evidencia-" & Format(i, "00") & "-" & Format(currentDate, "yyyy-mm-dd-hh-nn") & ".png"
        
        ' Exportar la diapositiva actual como PNG en la carpeta "Evidencias"
        sld.Export path & fileName, "PNG", 3333, 1875
        
        ' Incrementar el contador
        i = i + 1
    Next sld
    
    ' Mostrar mensaje de notificación con la ruta
    MsgBox "Se han guardado las evidencias en la siguiente ruta:" & vbNewLine & path, vbInformation
End Sub

