Attribute VB_Name = "M�dulo1"
Sub CambiarEstiloParrafoANormal()
    Dim parrafo As Paragraph
    
    For Each parrafo In ActiveDocument.Paragraphs
        If parrafo.Style = "P�rrafo" Then
            parrafo.Style = "Normal"
        End If
    Next parrafo
End Sub

