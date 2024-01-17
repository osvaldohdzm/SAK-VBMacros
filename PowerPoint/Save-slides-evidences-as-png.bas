Attribute VB_Name = "Módulo1"
Sub SaveEvidencesAsPNG()
    Dim sld As Slide
    Dim path As String
    Dim fileName As String
    Dim sPath As String
    Dim i As Integer
    Dim slideTitle As String
    Dim slideName As String
    Dim titleExists As Boolean
    Dim Shape As Shape
    
    ' Obtener la ubicación de la carpeta "Evidencias" en la ubicación de la presentación
    sPath = ActivePresentation.path
    path = sPath & "\Evidencias\"
    
    ' Comprobar si la carpeta "Evidencias" existe, si sí, borrarla
    If Dir(path, vbDirectory) <> "" Then
        On Error Resume Next ' Ignorar errores si la carpeta no se puede borrar
        Kill path & "*.*" ' Eliminar todos los archivos en la carpeta
        RmDir path ' Eliminar la carpeta
        On Error GoTo 0 ' Restablecer el manejo de errores
    End If
    
    ' Crear la carpeta "Evidencias" nuevamente
    MkDir path
    
    ' Inicializar el contador
    i = 1
    
    ' Recorrer todas las diapositivas en la presentación
    For Each sld In ActivePresentation.Slides
        ' Verificar si la diapositiva está marcada como oculta
        If Not sld.SlideShowTransition.Hidden Then
            ' La diapositiva no está oculta, continuamos con la exportación
            For Each Shape In sld.Shapes
                If Shape.Type = msoPicture Or Shape.Type = msoLinkedPicture Then
                    With Shape
                        .Shadow.Visible = True
                        .Shadow.ForeColor.RGB = RGB(0, 0, 0) ' Black color
                        .Shadow.Size = 102
                        .Shadow.Blur = 16
                        .Shadow.IncrementOffsetX 5
                        .Shadow.IncrementOffsetY 5
                        .Shadow.Transparency = 0.5
                        .Line.Visible = msoTrue
                        .Line.Weight = 1
                        .Line.Transparency = 0.5
                    End With
                End If
            Next Shape
            
            ' Verificar si la diapositiva tiene un título establecido
            titleExists = False
            If sld.Shapes.HasTitle Then
                slideTitle = sld.Shapes.Title.TextFrame.TextRange.Text
                If slideTitle <> "" Then
                    titleExists = True
                    ' Reemplazar caracteres no permitidos en nombres de archivo
                    slideName = Replace(slideTitle, ":", "")
                    slideName = Replace(slideName, "/", "")
                    slideName = Replace(slideName, "\", "")
                    slideName = Replace(slideName, "?", "")
                    slideName = Replace(slideName, "*", "")
                    slideName = Replace(slideName, " ", "-") ' Reemplazar espacios por guiones medios
                End If
            End If
            
            ' Formar el nombre del archivo utilizando el título o "Evidencia" y el contador
            If titleExists Then
                fileName = Format(i, "00") & "-" & slideName & ".png"
            Else
                fileName = Format(i, "00") & "-" & "Evidencia" & ".png"
            End If
            
            ' Exportar la diapositiva actual como PNG en la carpeta "Evidencias"
            sld.Export path & fileName, "PNG", 3333, 1875
            
            ' Incrementar el contador
            i = i + 1
        End If
    Next sld
    
    ' Mostrar mensaje de notificación con la ruta
    MsgBox "Se han guardado las evidencias en la siguiente ruta:" & vbNewLine & path, vbInformation
End Sub

