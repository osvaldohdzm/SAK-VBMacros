Attribute VB_Name = "Módulo1"
Sub CambiarEstiloParrafoANormal()
    Dim parrafo As Paragraph
    
    For Each parrafo In ActiveDocument.Paragraphs
        If parrafo.Style = "Párrafo" Then
            parrafo.Style = "Normal"
        End If
    Next parrafo
End Sub

